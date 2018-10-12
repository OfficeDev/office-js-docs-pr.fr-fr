
# <a name="item"></a><span data-ttu-id="7817f-101">item</span><span class="sxs-lookup"><span data-stu-id="7817f-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="7817f-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="7817f-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="7817f-p101">Utiliser l’espace-nom `item` pour accéder à vos messages , à vos demandes de réunion ou à vos rendez-vous. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="7817f-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="7817f-105">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-105">Requirements</span></span>

|<span data-ttu-id="7817f-106">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-106">Requirement</span></span>| <span data-ttu-id="7817f-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-108">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-109">1.0</span><span class="sxs-lookup"><span data-stu-id="7817f-109">1.0</span></span>|
|[<span data-ttu-id="7817f-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-111">Restreint</span><span class="sxs-lookup"><span data-stu-id="7817f-111">Restricted</span></span>|
|[<span data-ttu-id="7817f-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-113">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="7817f-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="7817f-114">Members and methods</span></span>

| <span data-ttu-id="7817f-115">Membre</span><span class="sxs-lookup"><span data-stu-id="7817f-115">Member</span></span> | <span data-ttu-id="7817f-116">Type</span><span class="sxs-lookup"><span data-stu-id="7817f-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="7817f-117">attachments</span><span class="sxs-lookup"><span data-stu-id="7817f-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails) | <span data-ttu-id="7817f-118">Membre</span><span class="sxs-lookup"><span data-stu-id="7817f-118">Member</span></span> |
| [<span data-ttu-id="7817f-119">bcc</span><span class="sxs-lookup"><span data-stu-id="7817f-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="7817f-120">Membre</span><span class="sxs-lookup"><span data-stu-id="7817f-120">Member</span></span> |
| [<span data-ttu-id="7817f-121">body</span><span class="sxs-lookup"><span data-stu-id="7817f-121">body</span></span>](#body-bodyjavascriptapioutlook16officebody) | <span data-ttu-id="7817f-122">Membre</span><span class="sxs-lookup"><span data-stu-id="7817f-122">Member</span></span> |
| [<span data-ttu-id="7817f-123">cc</span><span class="sxs-lookup"><span data-stu-id="7817f-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="7817f-124">Membre</span><span class="sxs-lookup"><span data-stu-id="7817f-124">Member</span></span> |
| [<span data-ttu-id="7817f-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="7817f-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="7817f-126">Membre</span><span class="sxs-lookup"><span data-stu-id="7817f-126">Member</span></span> |
| [<span data-ttu-id="7817f-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="7817f-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="7817f-128">Membre</span><span class="sxs-lookup"><span data-stu-id="7817f-128">Member</span></span> |
| [<span data-ttu-id="7817f-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="7817f-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="7817f-130">Membre</span><span class="sxs-lookup"><span data-stu-id="7817f-130">Member</span></span> |
| [<span data-ttu-id="7817f-131">end</span><span class="sxs-lookup"><span data-stu-id="7817f-131">end</span></span>](#end-datetimejavascriptapioutlook16officetime) | <span data-ttu-id="7817f-132">Membre</span><span class="sxs-lookup"><span data-stu-id="7817f-132">Member</span></span> |
| [<span data-ttu-id="7817f-133">from</span><span class="sxs-lookup"><span data-stu-id="7817f-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="7817f-134">Membre</span><span class="sxs-lookup"><span data-stu-id="7817f-134">Member</span></span> |
| [<span data-ttu-id="7817f-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="7817f-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="7817f-136">Membre</span><span class="sxs-lookup"><span data-stu-id="7817f-136">Member</span></span> |
| [<span data-ttu-id="7817f-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="7817f-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="7817f-138">Membre</span><span class="sxs-lookup"><span data-stu-id="7817f-138">Member</span></span> |
| [<span data-ttu-id="7817f-139">itemId</span><span class="sxs-lookup"><span data-stu-id="7817f-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="7817f-140">Membre</span><span class="sxs-lookup"><span data-stu-id="7817f-140">Member</span></span> |
| [<span data-ttu-id="7817f-141">itemType</span><span class="sxs-lookup"><span data-stu-id="7817f-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) | <span data-ttu-id="7817f-142">Membre</span><span class="sxs-lookup"><span data-stu-id="7817f-142">Member</span></span> |
| [<span data-ttu-id="7817f-143">location</span><span class="sxs-lookup"><span data-stu-id="7817f-143">location</span></span>](#location-stringlocationjavascriptapioutlook16officelocation) | <span data-ttu-id="7817f-144">Membre</span><span class="sxs-lookup"><span data-stu-id="7817f-144">Member</span></span> |
| [<span data-ttu-id="7817f-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="7817f-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="7817f-146">Membre</span><span class="sxs-lookup"><span data-stu-id="7817f-146">Member</span></span> |
| [<span data-ttu-id="7817f-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="7817f-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages) | <span data-ttu-id="7817f-148">Membre</span><span class="sxs-lookup"><span data-stu-id="7817f-148">Member</span></span> |
| [<span data-ttu-id="7817f-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="7817f-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="7817f-150">Membre</span><span class="sxs-lookup"><span data-stu-id="7817f-150">Member</span></span> |
| [<span data-ttu-id="7817f-151">organizer</span><span class="sxs-lookup"><span data-stu-id="7817f-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="7817f-152">Membre</span><span class="sxs-lookup"><span data-stu-id="7817f-152">Member</span></span> |
| [<span data-ttu-id="7817f-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="7817f-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="7817f-154">Membre</span><span class="sxs-lookup"><span data-stu-id="7817f-154">Member</span></span> |
| [<span data-ttu-id="7817f-155">sender</span><span class="sxs-lookup"><span data-stu-id="7817f-155">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="7817f-156">Membre</span><span class="sxs-lookup"><span data-stu-id="7817f-156">Member</span></span> |
| [<span data-ttu-id="7817f-157">start</span><span class="sxs-lookup"><span data-stu-id="7817f-157">start</span></span>](#start-datetimejavascriptapioutlook16officetime) | <span data-ttu-id="7817f-158">Membre</span><span class="sxs-lookup"><span data-stu-id="7817f-158">Member</span></span> |
| [<span data-ttu-id="7817f-159">subject</span><span class="sxs-lookup"><span data-stu-id="7817f-159">subject</span></span>](#subject-stringsubjectjavascriptapioutlook16officesubject) | <span data-ttu-id="7817f-160">Membre</span><span class="sxs-lookup"><span data-stu-id="7817f-160">Member</span></span> |
| [<span data-ttu-id="7817f-161">to</span><span class="sxs-lookup"><span data-stu-id="7817f-161">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="7817f-162">Membre</span><span class="sxs-lookup"><span data-stu-id="7817f-162">Member</span></span> |
| [<span data-ttu-id="7817f-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="7817f-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="7817f-164">Méthode</span><span class="sxs-lookup"><span data-stu-id="7817f-164">Method</span></span> |
| [<span data-ttu-id="7817f-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="7817f-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="7817f-166">Méthode</span><span class="sxs-lookup"><span data-stu-id="7817f-166">Method</span></span> |
| [<span data-ttu-id="7817f-167">close</span><span class="sxs-lookup"><span data-stu-id="7817f-167">close</span></span>](#close) | <span data-ttu-id="7817f-168">Méthode</span><span class="sxs-lookup"><span data-stu-id="7817f-168">Method</span></span> |
| [<span data-ttu-id="7817f-169">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="7817f-169">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="7817f-170">Méthode</span><span class="sxs-lookup"><span data-stu-id="7817f-170">Method</span></span> |
| [<span data-ttu-id="7817f-171">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="7817f-171">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="7817f-172">Méthode</span><span class="sxs-lookup"><span data-stu-id="7817f-172">Method</span></span> |
| [<span data-ttu-id="7817f-173">getEntities</span><span class="sxs-lookup"><span data-stu-id="7817f-173">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook16officeentities) | <span data-ttu-id="7817f-174">Méthode</span><span class="sxs-lookup"><span data-stu-id="7817f-174">Method</span></span> |
| [<span data-ttu-id="7817f-175">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="7817f-175">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | <span data-ttu-id="7817f-176">Méthode</span><span class="sxs-lookup"><span data-stu-id="7817f-176">Method</span></span> |
| [<span data-ttu-id="7817f-177">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="7817f-177">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | <span data-ttu-id="7817f-178">Méthode</span><span class="sxs-lookup"><span data-stu-id="7817f-178">Method</span></span> |
| [<span data-ttu-id="7817f-179">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="7817f-179">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="7817f-180">Méthode</span><span class="sxs-lookup"><span data-stu-id="7817f-180">Method</span></span> |
| [<span data-ttu-id="7817f-181">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="7817f-181">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="7817f-182">Méthode</span><span class="sxs-lookup"><span data-stu-id="7817f-182">Method</span></span> |
| [<span data-ttu-id="7817f-183">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="7817f-183">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="7817f-184">Méthode</span><span class="sxs-lookup"><span data-stu-id="7817f-184">Method</span></span> |
| [<span data-ttu-id="7817f-185">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="7817f-185">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlook16officeentities) | <span data-ttu-id="7817f-186">Méthode</span><span class="sxs-lookup"><span data-stu-id="7817f-186">Method</span></span> |
| [<span data-ttu-id="7817f-187">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="7817f-187">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="7817f-188">Méthode</span><span class="sxs-lookup"><span data-stu-id="7817f-188">Method</span></span> |
| [<span data-ttu-id="7817f-189">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="7817f-189">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="7817f-190">Méthode</span><span class="sxs-lookup"><span data-stu-id="7817f-190">Method</span></span> |
| [<span data-ttu-id="7817f-191">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="7817f-191">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="7817f-192">Méthode</span><span class="sxs-lookup"><span data-stu-id="7817f-192">Method</span></span> |
| [<span data-ttu-id="7817f-193">saveAsync</span><span class="sxs-lookup"><span data-stu-id="7817f-193">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="7817f-194">Méthode</span><span class="sxs-lookup"><span data-stu-id="7817f-194">Method</span></span> |
| [<span data-ttu-id="7817f-195">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="7817f-195">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="7817f-196">Méthode</span><span class="sxs-lookup"><span data-stu-id="7817f-196">Method</span></span> |

### <a name="example"></a><span data-ttu-id="7817f-197">Exemple</span><span class="sxs-lookup"><span data-stu-id="7817f-197">Example</span></span>

<span data-ttu-id="7817f-198">Cet exemple de code JavaScript montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="7817f-198">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```
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

### <a name="members"></a><span data-ttu-id="7817f-199">Membres</span><span class="sxs-lookup"><span data-stu-id="7817f-199">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails"></a><span data-ttu-id="7817f-200">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="7817f-200">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

<span data-ttu-id="7817f-p102">Obtient un tableau des pièces jointes de l’élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="7817f-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="7817f-203">Certains types de fichiers sont bloqués par Outlook en raison de problèmes de sécurité potentiels et ne sont donc pas rendus.</span><span class="sxs-lookup"><span data-stu-id="7817f-203">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="7817f-204">Pour plus d’information, voir les [pièces jointes bloquées par Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="7817f-204">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="7817f-205">Type :</span><span class="sxs-lookup"><span data-stu-id="7817f-205">Type:</span></span>

*   <span data-ttu-id="7817f-206">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="7817f-206">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="7817f-207">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-207">Requirements</span></span>

|<span data-ttu-id="7817f-208">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-208">Requirement</span></span>| <span data-ttu-id="7817f-209">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-210">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-210">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-211">1.0</span><span class="sxs-lookup"><span data-stu-id="7817f-211">1.0</span></span>|
|[<span data-ttu-id="7817f-212">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-212">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-213">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7817f-213">ReadItem</span></span>|
|[<span data-ttu-id="7817f-214">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-214">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-215">Lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-215">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7817f-216">Exemple</span><span class="sxs-lookup"><span data-stu-id="7817f-216">Example</span></span>

<span data-ttu-id="7817f-217">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="7817f-217">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```
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

####  <a name="bcc-recipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="7817f-218">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="7817f-218">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="7817f-219">Obtient un objet qui fournit les méthodes permettant d’obtenir ou de mettre à jour les destinataires des Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="7817f-219">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="7817f-220">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="7817f-220">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="7817f-221">Type :</span><span class="sxs-lookup"><span data-stu-id="7817f-221">Type:</span></span>

*   [<span data-ttu-id="7817f-222">Recipients</span><span class="sxs-lookup"><span data-stu-id="7817f-222">Recipients</span></span>](/javascript/api/outlook_1_6/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="7817f-223">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-223">Requirements</span></span>

|<span data-ttu-id="7817f-224">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-224">Requirement</span></span>| <span data-ttu-id="7817f-225">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-225">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-226">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-226">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-227">1.1</span><span class="sxs-lookup"><span data-stu-id="7817f-227">1.1</span></span>|
|[<span data-ttu-id="7817f-228">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-228">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-229">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7817f-229">ReadItem</span></span>|
|[<span data-ttu-id="7817f-230">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-230">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-231">Composition</span><span class="sxs-lookup"><span data-stu-id="7817f-231">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="7817f-232">Exemple</span><span class="sxs-lookup"><span data-stu-id="7817f-232">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook16officebody"></a><span data-ttu-id="7817f-233">body :[Body](/javascript/api/outlook_1_6/office.body)</span><span class="sxs-lookup"><span data-stu-id="7817f-233">body :[Body](/javascript/api/outlook_1_6/office.body)</span></span>

<span data-ttu-id="7817f-234">Obtient un objet qui fournit des méthodes permettant de manipuler le texte d’un élément.</span><span class="sxs-lookup"><span data-stu-id="7817f-234">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="7817f-235">Type :</span><span class="sxs-lookup"><span data-stu-id="7817f-235">Type:</span></span>

*   [<span data-ttu-id="7817f-236">Body</span><span class="sxs-lookup"><span data-stu-id="7817f-236">Body</span></span>](/javascript/api/outlook_1_6/office.body)

##### <a name="requirements"></a><span data-ttu-id="7817f-237">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-237">Requirements</span></span>

|<span data-ttu-id="7817f-238">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-238">Requirement</span></span>| <span data-ttu-id="7817f-239">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-240">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-240">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-241">1.1</span><span class="sxs-lookup"><span data-stu-id="7817f-241">1.1</span></span>|
|[<span data-ttu-id="7817f-242">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7817f-243">ReadItem</span></span>|
|[<span data-ttu-id="7817f-244">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-245">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-245">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="7817f-246">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="7817f-246">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="7817f-247">Permet d’accéder aux destinataires Cc (copie carbone) d’un message.</span><span class="sxs-lookup"><span data-stu-id="7817f-247">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="7817f-248">La nature d’un objet et son niveau d’accès dépendent du mode de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="7817f-248">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7817f-249">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-249">Read mode</span></span>

<span data-ttu-id="7817f-p106">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="7817f-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="7817f-252">Mode composition</span><span class="sxs-lookup"><span data-stu-id="7817f-252">Compose mode</span></span>

<span data-ttu-id="7817f-253">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="7817f-253">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="7817f-254">Type :</span><span class="sxs-lookup"><span data-stu-id="7817f-254">Type:</span></span>

*   <span data-ttu-id="7817f-255">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="7817f-255">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7817f-256">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-256">Requirements</span></span>

|<span data-ttu-id="7817f-257">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-257">Requirement</span></span>| <span data-ttu-id="7817f-258">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-259">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-259">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-260">1.0</span><span class="sxs-lookup"><span data-stu-id="7817f-260">1.0</span></span>|
|[<span data-ttu-id="7817f-261">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-261">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7817f-262">ReadItem</span></span>|
|[<span data-ttu-id="7817f-263">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-263">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-264">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-264">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="7817f-265">Exemple</span><span class="sxs-lookup"><span data-stu-id="7817f-265">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="7817f-266">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="7817f-266">(nullable) conversationId :String</span></span>

<span data-ttu-id="7817f-267">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="7817f-267">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="7817f-p107">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’identificateur de conversation de ce message changera et la valeur que vous aurez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="7817f-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="7817f-p108">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renverra une valeur.</span><span class="sxs-lookup"><span data-stu-id="7817f-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="7817f-272">Type :</span><span class="sxs-lookup"><span data-stu-id="7817f-272">Type:</span></span>

*   <span data-ttu-id="7817f-273">String</span><span class="sxs-lookup"><span data-stu-id="7817f-273">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7817f-274">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-274">Requirements</span></span>

|<span data-ttu-id="7817f-275">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-275">Requirement</span></span>| <span data-ttu-id="7817f-276">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-276">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-277">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-277">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-278">1.0</span><span class="sxs-lookup"><span data-stu-id="7817f-278">1.0</span></span>|
|[<span data-ttu-id="7817f-279">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-279">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-280">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7817f-280">ReadItem</span></span>|
|[<span data-ttu-id="7817f-281">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-281">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-282">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-282">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="7817f-283">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="7817f-283">dateTimeCreated :Date</span></span>

<span data-ttu-id="7817f-p109">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="7817f-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="7817f-286">Type :</span><span class="sxs-lookup"><span data-stu-id="7817f-286">Type:</span></span>

*   <span data-ttu-id="7817f-287">Date</span><span class="sxs-lookup"><span data-stu-id="7817f-287">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="7817f-288">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-288">Requirements</span></span>

|<span data-ttu-id="7817f-289">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-289">Requirement</span></span>| <span data-ttu-id="7817f-290">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-291">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-291">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-292">1.0</span><span class="sxs-lookup"><span data-stu-id="7817f-292">1.0</span></span>|
|[<span data-ttu-id="7817f-293">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-293">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7817f-294">ReadItem</span></span>|
|[<span data-ttu-id="7817f-295">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-295">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-296">Lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-296">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7817f-297">Exemple</span><span class="sxs-lookup"><span data-stu-id="7817f-297">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="7817f-298">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="7817f-298">dateTimeModified :Date</span></span>

<span data-ttu-id="7817f-p110">Obtient la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="7817f-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="7817f-301">Remarque : ce membre n’est pas pris en charge par Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="7817f-301">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="7817f-302">Type :</span><span class="sxs-lookup"><span data-stu-id="7817f-302">Type:</span></span>

*   <span data-ttu-id="7817f-303">Date</span><span class="sxs-lookup"><span data-stu-id="7817f-303">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="7817f-304">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-304">Requirements</span></span>

|<span data-ttu-id="7817f-305">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-305">Requirement</span></span>| <span data-ttu-id="7817f-306">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-307">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-307">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-308">1.0</span><span class="sxs-lookup"><span data-stu-id="7817f-308">1.0</span></span>|
|[<span data-ttu-id="7817f-309">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-309">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7817f-310">ReadItem</span></span>|
|[<span data-ttu-id="7817f-311">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-311">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-312">Lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7817f-313">Exemple</span><span class="sxs-lookup"><span data-stu-id="7817f-313">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="7817f-314">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="7817f-314">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="7817f-315">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="7817f-315">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="7817f-p111">La propriété `end` est exprimée en date et heure U.T.C. (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure locale du client.</span><span class="sxs-lookup"><span data-stu-id="7817f-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7817f-318">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-318">Read mode</span></span>

<span data-ttu-id="7817f-319">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="7817f-319">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="7817f-320">Mode composition</span><span class="sxs-lookup"><span data-stu-id="7817f-320">Compose mode</span></span>

<span data-ttu-id="7817f-321">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="7817f-321">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="7817f-322">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="7817f-322">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="7817f-323">Type :</span><span class="sxs-lookup"><span data-stu-id="7817f-323">Type:</span></span>

*   <span data-ttu-id="7817f-324">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="7817f-324">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7817f-325">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-325">Requirements</span></span>

|<span data-ttu-id="7817f-326">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-326">Requirement</span></span>| <span data-ttu-id="7817f-327">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-327">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-328">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-328">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-329">1.0</span><span class="sxs-lookup"><span data-stu-id="7817f-329">1.0</span></span>|
|[<span data-ttu-id="7817f-330">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-330">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-331">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7817f-331">ReadItem</span></span>|
|[<span data-ttu-id="7817f-332">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-332">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-333">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-333">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="7817f-334">Exemple</span><span class="sxs-lookup"><span data-stu-id="7817f-334">Example</span></span>

<span data-ttu-id="7817f-335">L’exemple suivant définit l’heure de fin d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="7817f-335">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```
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

#### <a name="from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="7817f-336">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="7817f-336">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="7817f-p112">Obtient l’adresse de messagerie de l’expéditeur d’un message. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="7817f-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="7817f-p113">Les propriétés `from` et [`sender`](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété expéditeur représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="7817f-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="7817f-341">La propriété  `recipientType` de l'objet  `EmailAddressDetails` dans la propriété  `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="7817f-341">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="7817f-342">Type :</span><span class="sxs-lookup"><span data-stu-id="7817f-342">Type:</span></span>

*   [<span data-ttu-id="7817f-343">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="7817f-343">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="7817f-344">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-344">Requirements</span></span>

|<span data-ttu-id="7817f-345">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-345">Requirement</span></span>| <span data-ttu-id="7817f-346">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-346">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-347">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-347">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-348">1.0</span><span class="sxs-lookup"><span data-stu-id="7817f-348">1.0</span></span>|
|[<span data-ttu-id="7817f-349">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-349">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-350">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7817f-350">ReadItem</span></span>|
|[<span data-ttu-id="7817f-351">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-351">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-352">Lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-352">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="7817f-353">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="7817f-353">internetMessageId :String</span></span>

<span data-ttu-id="7817f-p114">Obtient l’identificateur de message Internet d’un e-mail. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="7817f-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="7817f-356">Type :</span><span class="sxs-lookup"><span data-stu-id="7817f-356">Type:</span></span>

*   <span data-ttu-id="7817f-357">String</span><span class="sxs-lookup"><span data-stu-id="7817f-357">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7817f-358">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-358">Requirements</span></span>

|<span data-ttu-id="7817f-359">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-359">Requirement</span></span>| <span data-ttu-id="7817f-360">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-361">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-361">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-362">1.0</span><span class="sxs-lookup"><span data-stu-id="7817f-362">1.0</span></span>|
|[<span data-ttu-id="7817f-363">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-363">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7817f-364">ReadItem</span></span>|
|[<span data-ttu-id="7817f-365">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-365">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-366">Lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-366">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7817f-367">Exemple</span><span class="sxs-lookup"><span data-stu-id="7817f-367">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="7817f-368">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="7817f-368">itemClass :String</span></span>

<span data-ttu-id="7817f-p115">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="7817f-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="7817f-p116">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="7817f-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="7817f-373">Type</span><span class="sxs-lookup"><span data-stu-id="7817f-373">Type</span></span> | <span data-ttu-id="7817f-374">Description</span><span class="sxs-lookup"><span data-stu-id="7817f-374">Description</span></span> | <span data-ttu-id="7817f-375">Classe d’élément</span><span class="sxs-lookup"><span data-stu-id="7817f-375">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="7817f-376">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="7817f-376">Appointment items</span></span> | <span data-ttu-id="7817f-377">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="7817f-377">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="7817f-378">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="7817f-378">Message items</span></span> | <span data-ttu-id="7817f-379">Ces éléments incluent les e-mails dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="7817f-379">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="7817f-380">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="7817f-380">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="7817f-381">Type :</span><span class="sxs-lookup"><span data-stu-id="7817f-381">Type:</span></span>

*   <span data-ttu-id="7817f-382">String</span><span class="sxs-lookup"><span data-stu-id="7817f-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7817f-383">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-383">Requirements</span></span>

|<span data-ttu-id="7817f-384">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-384">Requirement</span></span>| <span data-ttu-id="7817f-385">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-386">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-386">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-387">1.0</span><span class="sxs-lookup"><span data-stu-id="7817f-387">1.0</span></span>|
|[<span data-ttu-id="7817f-388">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-388">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7817f-389">ReadItem</span></span>|
|[<span data-ttu-id="7817f-390">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-390">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-391">Lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7817f-392">Exemple</span><span class="sxs-lookup"><span data-stu-id="7817f-392">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="7817f-393">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="7817f-393">(nullable) itemId :String</span></span>

<span data-ttu-id="7817f-p117">Obtient l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="7817f-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="7817f-396">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="7817f-396">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="7817f-397">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ou l’ID utilisé par l’API REST de Outlook.</span><span class="sxs-lookup"><span data-stu-id="7817f-397">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="7817f-398">Avant d’effectuer des appels d’API REST à l’aide de cette valeur, elle doit être convertie à l’aide de [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="7817f-398">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="7817f-399">Pour plus d’informations, voir [Utiliser les API REST d’Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="7817f-399">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="7817f-p119">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="7817f-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="7817f-402">Type :</span><span class="sxs-lookup"><span data-stu-id="7817f-402">Type:</span></span>

*   <span data-ttu-id="7817f-403">String</span><span class="sxs-lookup"><span data-stu-id="7817f-403">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7817f-404">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-404">Requirements</span></span>

|<span data-ttu-id="7817f-405">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-405">Requirement</span></span>| <span data-ttu-id="7817f-406">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-406">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-407">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-407">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-408">1.0</span><span class="sxs-lookup"><span data-stu-id="7817f-408">1.0</span></span>|
|[<span data-ttu-id="7817f-409">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-409">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-410">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7817f-410">ReadItem</span></span>|
|[<span data-ttu-id="7817f-411">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-411">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-412">Lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-412">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7817f-413">Exemple</span><span class="sxs-lookup"><span data-stu-id="7817f-413">Example</span></span>

<span data-ttu-id="7817f-p120">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="7817f-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype"></a><span data-ttu-id="7817f-416">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="7817f-416">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="7817f-417">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="7817f-417">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="7817f-418">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="7817f-418">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="7817f-419">Type :</span><span class="sxs-lookup"><span data-stu-id="7817f-419">Type:</span></span>

*   [<span data-ttu-id="7817f-420">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="7817f-420">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="7817f-421">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-421">Requirements</span></span>

|<span data-ttu-id="7817f-422">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-422">Requirement</span></span>| <span data-ttu-id="7817f-423">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-424">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-424">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-425">1.0</span><span class="sxs-lookup"><span data-stu-id="7817f-425">1.0</span></span>|
|[<span data-ttu-id="7817f-426">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7817f-427">ReadItem</span></span>|
|[<span data-ttu-id="7817f-428">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-429">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-429">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="7817f-430">Exemple</span><span class="sxs-lookup"><span data-stu-id="7817f-430">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook16officelocation"></a><span data-ttu-id="7817f-431">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="7817f-431">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span></span>

<span data-ttu-id="7817f-432">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="7817f-432">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7817f-433">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-433">Read mode</span></span>

<span data-ttu-id="7817f-434">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="7817f-434">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="7817f-435">Mode composition</span><span class="sxs-lookup"><span data-stu-id="7817f-435">Compose mode</span></span>

<span data-ttu-id="7817f-436">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="7817f-436">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="7817f-437">Type :</span><span class="sxs-lookup"><span data-stu-id="7817f-437">Type:</span></span>

*   <span data-ttu-id="7817f-438">String | [Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="7817f-438">String | [Location](/javascript/api/outlook_1_6/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7817f-439">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-439">Requirements</span></span>

|<span data-ttu-id="7817f-440">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-440">Requirement</span></span>| <span data-ttu-id="7817f-441">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-442">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-442">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-443">1.0</span><span class="sxs-lookup"><span data-stu-id="7817f-443">1.0</span></span>|
|[<span data-ttu-id="7817f-444">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-444">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7817f-445">ReadItem</span></span>|
|[<span data-ttu-id="7817f-446">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-446">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-447">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-447">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="7817f-448">Exemple</span><span class="sxs-lookup"><span data-stu-id="7817f-448">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="7817f-449">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="7817f-449">normalizedSubject :String</span></span>

<span data-ttu-id="7817f-p121">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="7817f-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="7817f-p122">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubjectjavascriptapioutlook16officesubject).</span><span class="sxs-lookup"><span data-stu-id="7817f-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook16officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="7817f-454">Type :</span><span class="sxs-lookup"><span data-stu-id="7817f-454">Type:</span></span>

*   <span data-ttu-id="7817f-455">String</span><span class="sxs-lookup"><span data-stu-id="7817f-455">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7817f-456">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-456">Requirements</span></span>

|<span data-ttu-id="7817f-457">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-457">Requirement</span></span>| <span data-ttu-id="7817f-458">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-458">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-459">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-459">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-460">1.0</span><span class="sxs-lookup"><span data-stu-id="7817f-460">1.0</span></span>|
|[<span data-ttu-id="7817f-461">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-461">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-462">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7817f-462">ReadItem</span></span>|
|[<span data-ttu-id="7817f-463">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-463">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-464">Lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-464">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7817f-465">Exemple</span><span class="sxs-lookup"><span data-stu-id="7817f-465">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages"></a><span data-ttu-id="7817f-466">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="7817f-466">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span></span>

<span data-ttu-id="7817f-467">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="7817f-467">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="7817f-468">Type :</span><span class="sxs-lookup"><span data-stu-id="7817f-468">Type:</span></span>

*   [<span data-ttu-id="7817f-469">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="7817f-469">NotificationMessages</span></span>](/javascript/api/outlook_1_6/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="7817f-470">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-470">Requirements</span></span>

|<span data-ttu-id="7817f-471">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-471">Requirement</span></span>| <span data-ttu-id="7817f-472">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-472">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-473">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-473">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-474">1.3</span><span class="sxs-lookup"><span data-stu-id="7817f-474">1.3</span></span>|
|[<span data-ttu-id="7817f-475">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-475">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-476">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7817f-476">ReadItem</span></span>|
|[<span data-ttu-id="7817f-477">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-477">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-478">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-478">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="7817f-479">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="7817f-479">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="7817f-480">Fournit l’accès aux participants facultatifs d'un événement.</span><span class="sxs-lookup"><span data-stu-id="7817f-480">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="7817f-481">La nature d’un objet et son niveau d’accès dépendent du mode de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="7817f-481">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7817f-482">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-482">Read mode</span></span>

<span data-ttu-id="7817f-483">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="7817f-483">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="7817f-484">Mode composition</span><span class="sxs-lookup"><span data-stu-id="7817f-484">Compose mode</span></span>

<span data-ttu-id="7817f-485">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d'obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="7817f-485">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="7817f-486">Type :</span><span class="sxs-lookup"><span data-stu-id="7817f-486">Type:</span></span>

*   <span data-ttu-id="7817f-487">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="7817f-487">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7817f-488">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-488">Requirements</span></span>

|<span data-ttu-id="7817f-489">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-489">Requirement</span></span>| <span data-ttu-id="7817f-490">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-490">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-491">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-491">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-492">1.0</span><span class="sxs-lookup"><span data-stu-id="7817f-492">1.0</span></span>|
|[<span data-ttu-id="7817f-493">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-493">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-494">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7817f-494">ReadItem</span></span>|
|[<span data-ttu-id="7817f-495">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-495">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-496">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-496">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="7817f-497">Exemple</span><span class="sxs-lookup"><span data-stu-id="7817f-497">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="7817f-498">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="7817f-498">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="7817f-p124">Obtient l’adresse de messagerie de l’organisateur de la réunion spécifiée. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="7817f-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="7817f-501">Type :</span><span class="sxs-lookup"><span data-stu-id="7817f-501">Type:</span></span>

*   [<span data-ttu-id="7817f-502">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="7817f-502">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="7817f-503">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-503">Requirements</span></span>

|<span data-ttu-id="7817f-504">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-504">Requirement</span></span>| <span data-ttu-id="7817f-505">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-505">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-506">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-506">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-507">1.0</span><span class="sxs-lookup"><span data-stu-id="7817f-507">1.0</span></span>|
|[<span data-ttu-id="7817f-508">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-508">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7817f-509">ReadItem</span></span>|
|[<span data-ttu-id="7817f-510">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-510">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-511">Lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-511">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7817f-512">Exemple</span><span class="sxs-lookup"><span data-stu-id="7817f-512">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="7817f-513">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="7817f-513">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="7817f-514">Fournit l’accès aux participants obligatoires d'un événement.</span><span class="sxs-lookup"><span data-stu-id="7817f-514">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="7817f-515">La nature d’un objet et son niveau d’accès dépendent du mode de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="7817f-515">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7817f-516">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-516">Read mode</span></span>

<span data-ttu-id="7817f-517">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant obligatoires de la réunion.</span><span class="sxs-lookup"><span data-stu-id="7817f-517">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="7817f-518">Mode composition</span><span class="sxs-lookup"><span data-stu-id="7817f-518">Compose mode</span></span>

<span data-ttu-id="7817f-519">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d'obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="7817f-519">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="7817f-520">Type :</span><span class="sxs-lookup"><span data-stu-id="7817f-520">Type:</span></span>

*   <span data-ttu-id="7817f-521">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="7817f-521">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7817f-522">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-522">Requirements</span></span>

|<span data-ttu-id="7817f-523">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-523">Requirement</span></span>| <span data-ttu-id="7817f-524">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-524">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-525">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-525">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-526">1.0</span><span class="sxs-lookup"><span data-stu-id="7817f-526">1.0</span></span>|
|[<span data-ttu-id="7817f-527">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-527">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-528">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7817f-528">ReadItem</span></span>|
|[<span data-ttu-id="7817f-529">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-529">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-530">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-530">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="7817f-531">Exemple</span><span class="sxs-lookup"><span data-stu-id="7817f-531">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="7817f-532">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="7817f-532">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="7817f-p126">Obtient l’adresse de messagerie de l’expéditeur d’un e-mail. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="7817f-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="7817f-p127">Les propriétés [`from`](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété expéditeur représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="7817f-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="7817f-537">La propriété  `recipientType` de l'objet  `EmailAddressDetails` dans la propriété  `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="7817f-537">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="7817f-538">Type :</span><span class="sxs-lookup"><span data-stu-id="7817f-538">Type:</span></span>

*   [<span data-ttu-id="7817f-539">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="7817f-539">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="7817f-540">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-540">Requirements</span></span>

|<span data-ttu-id="7817f-541">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-541">Requirement</span></span>| <span data-ttu-id="7817f-542">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-542">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-543">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-543">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-544">1.0</span><span class="sxs-lookup"><span data-stu-id="7817f-544">1.0</span></span>|
|[<span data-ttu-id="7817f-545">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-545">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-546">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7817f-546">ReadItem</span></span>|
|[<span data-ttu-id="7817f-547">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-547">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-548">Lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-548">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7817f-549">Exemple</span><span class="sxs-lookup"><span data-stu-id="7817f-549">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="7817f-550">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="7817f-550">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="7817f-551">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="7817f-551">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="7817f-p128">La propriété `start` est exprimée en date et heure U.T.C. (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="7817f-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7817f-554">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-554">Read mode</span></span>

<span data-ttu-id="7817f-555">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="7817f-555">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="7817f-556">Mode composition</span><span class="sxs-lookup"><span data-stu-id="7817f-556">Compose mode</span></span>

<span data-ttu-id="7817f-557">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="7817f-557">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="7817f-558">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format U.T.C. pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="7817f-558">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="7817f-559">Type :</span><span class="sxs-lookup"><span data-stu-id="7817f-559">Type:</span></span>

*   <span data-ttu-id="7817f-560">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="7817f-560">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7817f-561">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-561">Requirements</span></span>

|<span data-ttu-id="7817f-562">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-562">Requirement</span></span>| <span data-ttu-id="7817f-563">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-564">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-564">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-565">1.0</span><span class="sxs-lookup"><span data-stu-id="7817f-565">1.0</span></span>|
|[<span data-ttu-id="7817f-566">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-567">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7817f-567">ReadItem</span></span>|
|[<span data-ttu-id="7817f-568">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-569">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-569">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="7817f-570">Exemple</span><span class="sxs-lookup"><span data-stu-id="7817f-570">Example</span></span>

<span data-ttu-id="7817f-571">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="7817f-571">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```
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

####  <a name="subject-stringsubjectjavascriptapioutlook16officesubject"></a><span data-ttu-id="7817f-572">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="7817f-572">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

<span data-ttu-id="7817f-573">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="7817f-573">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="7817f-574">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="7817f-574">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7817f-575">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-575">Read mode</span></span>

<span data-ttu-id="7817f-p129">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="7817f-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="7817f-578">Mode composition</span><span class="sxs-lookup"><span data-stu-id="7817f-578">Compose mode</span></span>

<span data-ttu-id="7817f-579">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="7817f-579">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="7817f-580">Type :</span><span class="sxs-lookup"><span data-stu-id="7817f-580">Type:</span></span>

*   <span data-ttu-id="7817f-581">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="7817f-581">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7817f-582">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-582">Requirements</span></span>

|<span data-ttu-id="7817f-583">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-583">Requirement</span></span>| <span data-ttu-id="7817f-584">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-584">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-585">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-585">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-586">1.0</span><span class="sxs-lookup"><span data-stu-id="7817f-586">1.0</span></span>|
|[<span data-ttu-id="7817f-587">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-587">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-588">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7817f-588">ReadItem</span></span>|
|[<span data-ttu-id="7817f-589">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-589">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-590">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-590">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="7817f-591">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="7817f-591">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="7817f-592">Permet d’accéder aux destinataires de la ligne **à** du message.</span><span class="sxs-lookup"><span data-stu-id="7817f-592">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="7817f-593">La nature d’un objet et son niveau d’accès dépendent du mode de l’élément actuel.</span><span class="sxs-lookup"><span data-stu-id="7817f-593">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="7817f-594">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-594">Read mode</span></span>

<span data-ttu-id="7817f-p131">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="7817f-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="7817f-597">Mode composition</span><span class="sxs-lookup"><span data-stu-id="7817f-597">Compose mode</span></span>

<span data-ttu-id="7817f-598">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="7817f-598">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="7817f-599">Type :</span><span class="sxs-lookup"><span data-stu-id="7817f-599">Type:</span></span>

*   <span data-ttu-id="7817f-600">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="7817f-600">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="7817f-601">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-601">Requirements</span></span>

|<span data-ttu-id="7817f-602">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-602">Requirement</span></span>| <span data-ttu-id="7817f-603">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-603">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-604">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-604">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-605">1.0</span><span class="sxs-lookup"><span data-stu-id="7817f-605">1.0</span></span>|
|[<span data-ttu-id="7817f-606">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-606">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-607">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7817f-607">ReadItem</span></span>|
|[<span data-ttu-id="7817f-608">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-608">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-609">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-609">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="7817f-610">Exemple</span><span class="sxs-lookup"><span data-stu-id="7817f-610">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="7817f-611">Méthodes</span><span class="sxs-lookup"><span data-stu-id="7817f-611">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="7817f-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="7817f-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="7817f-613">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="7817f-613">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="7817f-614">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="7817f-614">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="7817f-615">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="7817f-615">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7817f-616">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="7817f-616">Parameters:</span></span>

|<span data-ttu-id="7817f-617">Nom</span><span class="sxs-lookup"><span data-stu-id="7817f-617">Name</span></span>| <span data-ttu-id="7817f-618">Type</span><span class="sxs-lookup"><span data-stu-id="7817f-618">Type</span></span>| <span data-ttu-id="7817f-619">Attributs</span><span class="sxs-lookup"><span data-stu-id="7817f-619">Attributes</span></span>| <span data-ttu-id="7817f-620">Description</span><span class="sxs-lookup"><span data-stu-id="7817f-620">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="7817f-621">String</span><span class="sxs-lookup"><span data-stu-id="7817f-621">String</span></span>||<span data-ttu-id="7817f-p132">L’URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="7817f-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="7817f-624">String</span><span class="sxs-lookup"><span data-stu-id="7817f-624">String</span></span>||<span data-ttu-id="7817f-p133">Nom de la pièce jointe affiché lors de son chargement. La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="7817f-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="7817f-627">Object</span><span class="sxs-lookup"><span data-stu-id="7817f-627">Object</span></span>| <span data-ttu-id="7817f-628">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="7817f-628">&lt;optional&gt;</span></span>|<span data-ttu-id="7817f-629">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="7817f-629">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="7817f-630">Object</span><span class="sxs-lookup"><span data-stu-id="7817f-630">Object</span></span> | <span data-ttu-id="7817f-631">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="7817f-631">&lt;optional&gt;</span></span> | <span data-ttu-id="7817f-632">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="7817f-632">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="7817f-633">Boolean</span><span class="sxs-lookup"><span data-stu-id="7817f-633">Boolean</span></span> | <span data-ttu-id="7817f-634">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="7817f-634">&lt;optional&gt;</span></span> | <span data-ttu-id="7817f-635">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le texte du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="7817f-635">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="7817f-636">fonction</span><span class="sxs-lookup"><span data-stu-id="7817f-636">function</span></span>| <span data-ttu-id="7817f-637">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="7817f-637">&lt;optional&gt;</span></span>|<span data-ttu-id="7817f-638">Quand la méthode se termine, la fonction passée dans le paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7817f-638">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="7817f-639">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="7817f-639">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="7817f-640">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="7817f-640">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="7817f-641">Erreurs</span><span class="sxs-lookup"><span data-stu-id="7817f-641">Errors</span></span>

| <span data-ttu-id="7817f-642">Code d’erreur</span><span class="sxs-lookup"><span data-stu-id="7817f-642">Error code</span></span> | <span data-ttu-id="7817f-643">Description</span><span class="sxs-lookup"><span data-stu-id="7817f-643">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="7817f-644">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="7817f-644">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="7817f-645">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="7817f-645">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="7817f-646">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="7817f-646">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7817f-647">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-647">Requirements</span></span>

|<span data-ttu-id="7817f-648">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-648">Requirement</span></span>| <span data-ttu-id="7817f-649">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-650">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-650">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-651">1.1</span><span class="sxs-lookup"><span data-stu-id="7817f-651">1.1</span></span>|
|[<span data-ttu-id="7817f-652">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-652">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-653">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="7817f-653">ReadWriteItem</span></span>|
|[<span data-ttu-id="7817f-654">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-654">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-655">Composition</span><span class="sxs-lookup"><span data-stu-id="7817f-655">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="7817f-656">Exemples</span><span class="sxs-lookup"><span data-stu-id="7817f-656">Examples</span></span>

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

<span data-ttu-id="7817f-657">L’exemple suivant montre comment ajouter un fichier image comme pièce jointe incorporée et comment la pièce jointe est affichée dans le texte du message.</span><span class="sxs-lookup"><span data-stu-id="7817f-657">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="7817f-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="7817f-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="7817f-659">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="7817f-659">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="7817f-p134">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="7817f-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="7817f-663">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="7817f-663">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="7817f-664">Si votre complément Office est exécuté dans la Outlook Web App , la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="7817f-664">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7817f-665">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="7817f-665">Parameters:</span></span>

|<span data-ttu-id="7817f-666">Nom</span><span class="sxs-lookup"><span data-stu-id="7817f-666">Name</span></span>| <span data-ttu-id="7817f-667">Type</span><span class="sxs-lookup"><span data-stu-id="7817f-667">Type</span></span>| <span data-ttu-id="7817f-668">Attributs</span><span class="sxs-lookup"><span data-stu-id="7817f-668">Attributes</span></span>| <span data-ttu-id="7817f-669">Description</span><span class="sxs-lookup"><span data-stu-id="7817f-669">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="7817f-670">String</span><span class="sxs-lookup"><span data-stu-id="7817f-670">String</span></span>||<span data-ttu-id="7817f-p135">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="7817f-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="7817f-673">String</span><span class="sxs-lookup"><span data-stu-id="7817f-673">String</span></span>||<span data-ttu-id="7817f-p136">Objet de l’élément à joindre. La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="7817f-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="7817f-676">Object</span><span class="sxs-lookup"><span data-stu-id="7817f-676">Object</span></span>| <span data-ttu-id="7817f-677">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="7817f-677">&lt;optional&gt;</span></span>|<span data-ttu-id="7817f-678">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="7817f-678">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="7817f-679">Object</span><span class="sxs-lookup"><span data-stu-id="7817f-679">Object</span></span>| <span data-ttu-id="7817f-680">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="7817f-680">&lt;optional&gt;</span></span>|<span data-ttu-id="7817f-681">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="7817f-681">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="7817f-682">fonction</span><span class="sxs-lookup"><span data-stu-id="7817f-682">function</span></span>| <span data-ttu-id="7817f-683">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="7817f-683">&lt;optional&gt;</span></span>|<span data-ttu-id="7817f-684">Quand la méthode se termine, la fonction passée dans le paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7817f-684">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="7817f-685">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="7817f-685">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="7817f-686">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="7817f-686">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="7817f-687">Erreurs</span><span class="sxs-lookup"><span data-stu-id="7817f-687">Errors</span></span>

| <span data-ttu-id="7817f-688">Code d’erreur</span><span class="sxs-lookup"><span data-stu-id="7817f-688">Error code</span></span> | <span data-ttu-id="7817f-689">Description</span><span class="sxs-lookup"><span data-stu-id="7817f-689">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="7817f-690">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="7817f-690">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7817f-691">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-691">Requirements</span></span>

|<span data-ttu-id="7817f-692">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-692">Requirement</span></span>| <span data-ttu-id="7817f-693">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-693">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-694">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-694">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-695">1.1</span><span class="sxs-lookup"><span data-stu-id="7817f-695">1.1</span></span>|
|[<span data-ttu-id="7817f-696">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-696">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-697">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="7817f-697">ReadWriteItem</span></span>|
|[<span data-ttu-id="7817f-698">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-698">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-699">Composition</span><span class="sxs-lookup"><span data-stu-id="7817f-699">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="7817f-700">Exemple</span><span class="sxs-lookup"><span data-stu-id="7817f-700">Example</span></span>

<span data-ttu-id="7817f-701">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="7817f-701">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```
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

####  <a name="close"></a><span data-ttu-id="7817f-702">close()</span><span class="sxs-lookup"><span data-stu-id="7817f-702">close()</span></span>

<span data-ttu-id="7817f-703">Ferme l’élément actif qui est en train d’être composé.</span><span class="sxs-lookup"><span data-stu-id="7817f-703">Closes the current item that is being composed.</span></span>

<span data-ttu-id="7817f-p137">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action fermer.</span><span class="sxs-lookup"><span data-stu-id="7817f-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="7817f-706">Sur Outlook Web Access, si l’élément est un rendez-vous qui a déjà été sauvegardé en utilisant la méthode `saveAsync` , l'utilisateur sera inviter à sauvegarder, abandonner ou annuler même si l’élément n'a subi aucun changement depuis sa dernière sauvegarde.</span><span class="sxs-lookup"><span data-stu-id="7817f-706">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="7817f-707">Dans la version logicielle d'Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="7817f-707">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="7817f-708">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-708">Requirements</span></span>

|<span data-ttu-id="7817f-709">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-709">Requirement</span></span>| <span data-ttu-id="7817f-710">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-710">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-711">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-711">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-712">1.3</span><span class="sxs-lookup"><span data-stu-id="7817f-712">1.3</span></span>|
|[<span data-ttu-id="7817f-713">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-713">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-714">Restreint</span><span class="sxs-lookup"><span data-stu-id="7817f-714">Restricted</span></span>|
|[<span data-ttu-id="7817f-715">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-715">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-716">Composition</span><span class="sxs-lookup"><span data-stu-id="7817f-716">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="7817f-717">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="7817f-717">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="7817f-718">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="7817f-718">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="7817f-719">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="7817f-719">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="7817f-720">Sur Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="7817f-720">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="7817f-721">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="7817f-721">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="7817f-p138">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, alors aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="7817f-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7817f-725">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="7817f-725">Parameters:</span></span>

| <span data-ttu-id="7817f-726">Nom</span><span class="sxs-lookup"><span data-stu-id="7817f-726">Name</span></span> | <span data-ttu-id="7817f-727">Type</span><span class="sxs-lookup"><span data-stu-id="7817f-727">Type</span></span> | <span data-ttu-id="7817f-728">Attributs</span><span class="sxs-lookup"><span data-stu-id="7817f-728">Attributes</span></span> | <span data-ttu-id="7817f-729">Description</span><span class="sxs-lookup"><span data-stu-id="7817f-729">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="7817f-730">String | Object</span><span class="sxs-lookup"><span data-stu-id="7817f-730">String &#124; Object</span></span>| |<span data-ttu-id="7817f-p139">Chaîne qui contient du texte et des éléments HTML et qui représente le texte du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="7817f-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="7817f-733">**OU**</span><span class="sxs-lookup"><span data-stu-id="7817f-733">**OR**</span></span><br/><span data-ttu-id="7817f-p140">Objet qui contient les données du texte du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="7817f-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="7817f-736">String</span><span class="sxs-lookup"><span data-stu-id="7817f-736">String</span></span> | <span data-ttu-id="7817f-737">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="7817f-737">&lt;optional&gt;</span></span> | <span data-ttu-id="7817f-p141">Chaîne qui contient du texte et des éléments HTML et qui représente le texte du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="7817f-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="7817f-740">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="7817f-740">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="7817f-741">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="7817f-741">&lt;optional&gt;</span></span> | <span data-ttu-id="7817f-742">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="7817f-742">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="7817f-743">String</span><span class="sxs-lookup"><span data-stu-id="7817f-743">String</span></span> | | <span data-ttu-id="7817f-p142">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="7817f-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="7817f-746">String</span><span class="sxs-lookup"><span data-stu-id="7817f-746">String</span></span> | | <span data-ttu-id="7817f-747">Chaîne qui contient le nom de la pièce jointe, d'une longueur maximale de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="7817f-747">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="7817f-748">String</span><span class="sxs-lookup"><span data-stu-id="7817f-748">String</span></span> | | <span data-ttu-id="7817f-p143">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="7817f-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="7817f-751">Boolean</span><span class="sxs-lookup"><span data-stu-id="7817f-751">Boolean</span></span> | | <span data-ttu-id="7817f-p144">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le texte du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="7817f-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="7817f-754">String</span><span class="sxs-lookup"><span data-stu-id="7817f-754">String</span></span> | | <span data-ttu-id="7817f-p145">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’identificateur de l’élément EWS de la pièce jointe. Cette chaîne doit être d'une longueur maximale de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="7817f-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="7817f-758">fonction</span><span class="sxs-lookup"><span data-stu-id="7817f-758">function</span></span> | <span data-ttu-id="7817f-759">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="7817f-759">&lt;optional&gt;</span></span> | <span data-ttu-id="7817f-760">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7817f-760">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7817f-761">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-761">Requirements</span></span>

|<span data-ttu-id="7817f-762">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-762">Requirement</span></span>| <span data-ttu-id="7817f-763">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-763">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-764">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-764">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-765">1.0</span><span class="sxs-lookup"><span data-stu-id="7817f-765">1.0</span></span>|
|[<span data-ttu-id="7817f-766">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-766">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-767">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7817f-767">ReadItem</span></span>|
|[<span data-ttu-id="7817f-768">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-768">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-769">Lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-769">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="7817f-770">Exemples</span><span class="sxs-lookup"><span data-stu-id="7817f-770">Examples</span></span>

<span data-ttu-id="7817f-771">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="7817f-771">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="7817f-772">Réponse sans texte.</span><span class="sxs-lookup"><span data-stu-id="7817f-772">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="7817f-773">Réponse avec un texte.</span><span class="sxs-lookup"><span data-stu-id="7817f-773">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="7817f-774">Réponse avec un texte et un fichier comme pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="7817f-774">Reply with a body and a file attachment.</span></span>

```
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

<span data-ttu-id="7817f-775">Réponse avec un texte et un élément comme pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="7817f-775">Reply with a body and an item attachment.</span></span>

```
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

<span data-ttu-id="7817f-776">Réponse avec un texte, un fichier comme pièce jointe, un élément comme pièce jointe et un rappel.</span><span class="sxs-lookup"><span data-stu-id="7817f-776">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```
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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="7817f-777">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="7817f-777">displayReplyForm(formData)</span></span>

<span data-ttu-id="7817f-778">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="7817f-778">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="7817f-779">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="7817f-779">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="7817f-780">Sur Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="7817f-780">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="7817f-781">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="7817f-781">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="7817f-p146">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, alors aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="7817f-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7817f-785">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="7817f-785">Parameters:</span></span>

| <span data-ttu-id="7817f-786">Nom</span><span class="sxs-lookup"><span data-stu-id="7817f-786">Name</span></span> | <span data-ttu-id="7817f-787">Type</span><span class="sxs-lookup"><span data-stu-id="7817f-787">Type</span></span> | <span data-ttu-id="7817f-788">Attributs</span><span class="sxs-lookup"><span data-stu-id="7817f-788">Attributes</span></span> | <span data-ttu-id="7817f-789">Description</span><span class="sxs-lookup"><span data-stu-id="7817f-789">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="7817f-790">String | Object</span><span class="sxs-lookup"><span data-stu-id="7817f-790">String &#124; Object</span></span>| | <span data-ttu-id="7817f-p147">Chaîne qui contient du texte et des éléments HTML et qui représente le texte du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="7817f-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="7817f-793">**OU**</span><span class="sxs-lookup"><span data-stu-id="7817f-793">**OR**</span></span><br/><span data-ttu-id="7817f-p148">Objet qui contient les données du texte du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="7817f-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="7817f-796">String</span><span class="sxs-lookup"><span data-stu-id="7817f-796">String</span></span> | <span data-ttu-id="7817f-797">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="7817f-797">&lt;optional&gt;</span></span> | <span data-ttu-id="7817f-p149">Chaîne qui contient du texte et des éléments HTML et qui représente le texte du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="7817f-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="7817f-800">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="7817f-800">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="7817f-801">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="7817f-801">&lt;optional&gt;</span></span> | <span data-ttu-id="7817f-802">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="7817f-802">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="7817f-803">String</span><span class="sxs-lookup"><span data-stu-id="7817f-803">String</span></span> | | <span data-ttu-id="7817f-p150">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="7817f-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="7817f-806">String</span><span class="sxs-lookup"><span data-stu-id="7817f-806">String</span></span> | | <span data-ttu-id="7817f-807">Chaîne qui contient le nom de la pièce jointe, d'une longueur maximale de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="7817f-807">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="7817f-808">String</span><span class="sxs-lookup"><span data-stu-id="7817f-808">String</span></span> | | <span data-ttu-id="7817f-p151">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="7817f-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="7817f-811">Boolean</span><span class="sxs-lookup"><span data-stu-id="7817f-811">Boolean</span></span> | | <span data-ttu-id="7817f-p152">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le texte du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="7817f-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="7817f-814">String</span><span class="sxs-lookup"><span data-stu-id="7817f-814">String</span></span> | | <span data-ttu-id="7817f-p153">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’identificateur de l’élément EWS de la pièce jointe. Cette chaîne doit être d'une longueur maximale de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="7817f-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="7817f-818">fonction</span><span class="sxs-lookup"><span data-stu-id="7817f-818">function</span></span> | <span data-ttu-id="7817f-819">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="7817f-819">&lt;optional&gt;</span></span> | <span data-ttu-id="7817f-820">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7817f-820">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7817f-821">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-821">Requirements</span></span>

|<span data-ttu-id="7817f-822">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-822">Requirement</span></span>| <span data-ttu-id="7817f-823">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-823">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-824">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-824">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-825">1.0</span><span class="sxs-lookup"><span data-stu-id="7817f-825">1.0</span></span>|
|[<span data-ttu-id="7817f-826">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-826">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-827">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7817f-827">ReadItem</span></span>|
|[<span data-ttu-id="7817f-828">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-828">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-829">Lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-829">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="7817f-830">Exemples</span><span class="sxs-lookup"><span data-stu-id="7817f-830">Examples</span></span>

<span data-ttu-id="7817f-831">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="7817f-831">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="7817f-832">Réponse sans texte.</span><span class="sxs-lookup"><span data-stu-id="7817f-832">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="7817f-833">Réponse avec un texte.</span><span class="sxs-lookup"><span data-stu-id="7817f-833">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="7817f-834">Réponse avec un texte et un fichier comme pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="7817f-834">Reply with a body and a file attachment.</span></span>

```
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

<span data-ttu-id="7817f-835">Réponse avec un texte et un élément comme pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="7817f-835">Reply with a body and an item attachment.</span></span>

```
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

<span data-ttu-id="7817f-836">Réponse avec un texte, un fichier comme pièce jointe, un élément comme pièce jointe et un rappel.</span><span class="sxs-lookup"><span data-stu-id="7817f-836">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```
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

#### <a name="getentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="7817f-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="7817f-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="7817f-838">Obtient les entités figurant dans le texte de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="7817f-838">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="7817f-839">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="7817f-839">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="7817f-840">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-840">Requirements</span></span>

|<span data-ttu-id="7817f-841">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-841">Requirement</span></span>| <span data-ttu-id="7817f-842">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-842">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-843">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-843">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-844">1.0</span><span class="sxs-lookup"><span data-stu-id="7817f-844">1.0</span></span>|
|[<span data-ttu-id="7817f-845">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-845">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-846">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7817f-846">ReadItem</span></span>|
|[<span data-ttu-id="7817f-847">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-847">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-848">Lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-848">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7817f-849">Retourne :</span><span class="sxs-lookup"><span data-stu-id="7817f-849">Returns:</span></span>

<span data-ttu-id="7817f-850">Type : [Entités](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="7817f-850">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="7817f-851">Exemple</span><span class="sxs-lookup"><span data-stu-id="7817f-851">Example</span></span>

<span data-ttu-id="7817f-852">L’exemple suivant accède aux entités de contacts dans l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="7817f-852">The following example accesses the contacts entities on the current item.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="7817f-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="7817f-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="7817f-854">Obtient un tableau de toutes les entités du type spécifié trouvées dans le texte de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="7817f-854">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="7817f-855">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="7817f-855">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7817f-856">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="7817f-856">Parameters:</span></span>

|<span data-ttu-id="7817f-857">Nom</span><span class="sxs-lookup"><span data-stu-id="7817f-857">Name</span></span>| <span data-ttu-id="7817f-858">Type</span><span class="sxs-lookup"><span data-stu-id="7817f-858">Type</span></span>| <span data-ttu-id="7817f-859">Description</span><span class="sxs-lookup"><span data-stu-id="7817f-859">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="7817f-860">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="7817f-860">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.entitytype)|<span data-ttu-id="7817f-861">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="7817f-861">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7817f-862">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-862">Requirements</span></span>

|<span data-ttu-id="7817f-863">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-863">Requirement</span></span>| <span data-ttu-id="7817f-864">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-864">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-865">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-865">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-866">1.0</span><span class="sxs-lookup"><span data-stu-id="7817f-866">1.0</span></span>|
|[<span data-ttu-id="7817f-867">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-867">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-868">Restreint</span><span class="sxs-lookup"><span data-stu-id="7817f-868">Restricted</span></span>|
|[<span data-ttu-id="7817f-869">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-869">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-870">Lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-870">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7817f-871">Retourne :</span><span class="sxs-lookup"><span data-stu-id="7817f-871">Returns:</span></span>

<span data-ttu-id="7817f-872">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="7817f-872">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="7817f-873">Si aucune entité du type spécifié n’est présente dans le texte de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="7817f-873">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="7817f-874">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="7817f-874">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="7817f-875">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="7817f-875">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="7817f-876">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="7817f-876">Value of `entityType`</span></span> | <span data-ttu-id="7817f-877">Type des objets dans le tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="7817f-877">Type of objects in returned array</span></span> | <span data-ttu-id="7817f-878">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="7817f-878">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="7817f-879">String</span><span class="sxs-lookup"><span data-stu-id="7817f-879">String</span></span> | <span data-ttu-id="7817f-880">**Restreint**</span><span class="sxs-lookup"><span data-stu-id="7817f-880">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="7817f-881">Contact</span><span class="sxs-lookup"><span data-stu-id="7817f-881">Contact</span></span> | <span data-ttu-id="7817f-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="7817f-882">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="7817f-883">String</span><span class="sxs-lookup"><span data-stu-id="7817f-883">String</span></span> | <span data-ttu-id="7817f-884">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="7817f-884">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="7817f-885">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="7817f-885">MeetingSuggestion</span></span> | <span data-ttu-id="7817f-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="7817f-886">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="7817f-887">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="7817f-887">PhoneNumber</span></span> | <span data-ttu-id="7817f-888">**Restreint**</span><span class="sxs-lookup"><span data-stu-id="7817f-888">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="7817f-889">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="7817f-889">TaskSuggestion</span></span> | <span data-ttu-id="7817f-890">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="7817f-890">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="7817f-891">String</span><span class="sxs-lookup"><span data-stu-id="7817f-891">String</span></span> | <span data-ttu-id="7817f-892">**Restreint**</span><span class="sxs-lookup"><span data-stu-id="7817f-892">**Restricted**</span></span> |

<span data-ttu-id="7817f-893">Type : Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="7817f-893">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="7817f-894">Exemple</span><span class="sxs-lookup"><span data-stu-id="7817f-894">Example</span></span>

<span data-ttu-id="7817f-895">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le texte de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="7817f-895">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

```
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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="7817f-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="7817f-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="7817f-897">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="7817f-897">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="7817f-898">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="7817f-898">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="7817f-899">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="7817f-899">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7817f-900">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="7817f-900">Parameters:</span></span>

|<span data-ttu-id="7817f-901">Nom</span><span class="sxs-lookup"><span data-stu-id="7817f-901">Name</span></span>| <span data-ttu-id="7817f-902">Type</span><span class="sxs-lookup"><span data-stu-id="7817f-902">Type</span></span>| <span data-ttu-id="7817f-903">Description</span><span class="sxs-lookup"><span data-stu-id="7817f-903">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="7817f-904">String</span><span class="sxs-lookup"><span data-stu-id="7817f-904">String</span></span>|<span data-ttu-id="7817f-905">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="7817f-905">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7817f-906">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-906">Requirements</span></span>

|<span data-ttu-id="7817f-907">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-907">Requirement</span></span>| <span data-ttu-id="7817f-908">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-908">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-909">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-909">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-910">1.0</span><span class="sxs-lookup"><span data-stu-id="7817f-910">1.0</span></span>|
|[<span data-ttu-id="7817f-911">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-911">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-912">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7817f-912">ReadItem</span></span>|
|[<span data-ttu-id="7817f-913">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-913">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-914">Lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-914">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7817f-915">Retourne :</span><span class="sxs-lookup"><span data-stu-id="7817f-915">Returns:</span></span>

<span data-ttu-id="7817f-p155">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="7817f-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="7817f-918">Type : Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="7817f-918">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="7817f-919">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="7817f-919">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="7817f-920">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="7817f-920">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="7817f-921">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="7817f-921">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="7817f-p156">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="7817f-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="7817f-925">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="7817f-925">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="7817f-926">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="7817f-926">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="7817f-p157">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le texte. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du texte de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du texte d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du texte de l’élément.</span><span class="sxs-lookup"><span data-stu-id="7817f-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="7817f-930">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-930">Requirements</span></span>

|<span data-ttu-id="7817f-931">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-931">Requirement</span></span>| <span data-ttu-id="7817f-932">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-932">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-933">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-933">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-934">1.0</span><span class="sxs-lookup"><span data-stu-id="7817f-934">1.0</span></span>|
|[<span data-ttu-id="7817f-935">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-935">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-936">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7817f-936">ReadItem</span></span>|
|[<span data-ttu-id="7817f-937">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-937">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-938">Lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-938">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7817f-939">Retourne :</span><span class="sxs-lookup"><span data-stu-id="7817f-939">Returns:</span></span>

<span data-ttu-id="7817f-p158">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="7817f-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="7817f-942">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="7817f-942">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="7817f-943">Object</span><span class="sxs-lookup"><span data-stu-id="7817f-943">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="7817f-944">Exemple</span><span class="sxs-lookup"><span data-stu-id="7817f-944">Example</span></span>

<span data-ttu-id="7817f-945">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="7817f-945">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="7817f-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="7817f-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="7817f-947">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="7817f-947">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="7817f-948">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="7817f-948">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="7817f-949">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="7817f-949">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="7817f-p159">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de texte d’un élément, l’expression régulière doit filtrer davantage le texte. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du texte de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du texte d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="7817f-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7817f-952">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="7817f-952">Parameters:</span></span>

|<span data-ttu-id="7817f-953">Nom</span><span class="sxs-lookup"><span data-stu-id="7817f-953">Name</span></span>| <span data-ttu-id="7817f-954">Type</span><span class="sxs-lookup"><span data-stu-id="7817f-954">Type</span></span>| <span data-ttu-id="7817f-955">Description</span><span class="sxs-lookup"><span data-stu-id="7817f-955">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="7817f-956">String</span><span class="sxs-lookup"><span data-stu-id="7817f-956">String</span></span>|<span data-ttu-id="7817f-957">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="7817f-957">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7817f-958">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-958">Requirements</span></span>

|<span data-ttu-id="7817f-959">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-959">Requirement</span></span>| <span data-ttu-id="7817f-960">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-960">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-961">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-961">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-962">1.0</span><span class="sxs-lookup"><span data-stu-id="7817f-962">1.0</span></span>|
|[<span data-ttu-id="7817f-963">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-963">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-964">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7817f-964">ReadItem</span></span>|
|[<span data-ttu-id="7817f-965">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-965">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-966">Lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-966">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7817f-967">Retourne :</span><span class="sxs-lookup"><span data-stu-id="7817f-967">Returns:</span></span>

<span data-ttu-id="7817f-968">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="7817f-968">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="7817f-969">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="7817f-969">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="7817f-970">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="7817f-970">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="7817f-971">Exemple</span><span class="sxs-lookup"><span data-stu-id="7817f-971">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="7817f-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="7817f-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="7817f-973">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du texte d’un message.</span><span class="sxs-lookup"><span data-stu-id="7817f-973">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="7817f-p160">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="7817f-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7817f-976">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="7817f-976">Parameters:</span></span>

|<span data-ttu-id="7817f-977">Nom</span><span class="sxs-lookup"><span data-stu-id="7817f-977">Name</span></span>| <span data-ttu-id="7817f-978">Type</span><span class="sxs-lookup"><span data-stu-id="7817f-978">Type</span></span>| <span data-ttu-id="7817f-979">Attributs</span><span class="sxs-lookup"><span data-stu-id="7817f-979">Attributes</span></span>| <span data-ttu-id="7817f-980">Description</span><span class="sxs-lookup"><span data-stu-id="7817f-980">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="7817f-981">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="7817f-981">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="7817f-p161">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="7817f-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="7817f-985">Object</span><span class="sxs-lookup"><span data-stu-id="7817f-985">Object</span></span>| <span data-ttu-id="7817f-986">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="7817f-986">&lt;optional&gt;</span></span>|<span data-ttu-id="7817f-987">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="7817f-987">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="7817f-988">Object</span><span class="sxs-lookup"><span data-stu-id="7817f-988">Object</span></span>| <span data-ttu-id="7817f-989">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="7817f-989">&lt;optional&gt;</span></span>|<span data-ttu-id="7817f-990">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="7817f-990">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="7817f-991">fonction</span><span class="sxs-lookup"><span data-stu-id="7817f-991">function</span></span>||<span data-ttu-id="7817f-992">Quand la méthode se termine, la fonction passée dans le paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7817f-992">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="7817f-993">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="7817f-993">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="7817f-994">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`  , qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="7817f-994">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7817f-995">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-995">Requirements</span></span>

|<span data-ttu-id="7817f-996">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-996">Requirement</span></span>| <span data-ttu-id="7817f-997">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-997">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-998">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-998">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-999">1.2</span><span class="sxs-lookup"><span data-stu-id="7817f-999">1.2</span></span>|
|[<span data-ttu-id="7817f-1000">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-1000">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-1001">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="7817f-1001">ReadWriteItem</span></span>|
|[<span data-ttu-id="7817f-1002">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-1002">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-1003">Composition</span><span class="sxs-lookup"><span data-stu-id="7817f-1003">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="7817f-1004">Retourne :</span><span class="sxs-lookup"><span data-stu-id="7817f-1004">Returns:</span></span>

<span data-ttu-id="7817f-1005">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="7817f-1005">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="7817f-1006">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="7817f-1006">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="7817f-1007">String</span><span class="sxs-lookup"><span data-stu-id="7817f-1007">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="7817f-1008">Exemple</span><span class="sxs-lookup"><span data-stu-id="7817f-1008">Example</span></span>

```
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

#### <a name="getselectedentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="7817f-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="7817f-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="7817f-p163">Obtient les entités figurant dans une correspondance en surbrillance qu’un utilisateur a sélectionné. Les correspondances en surbrillance s’appliquent aux [compléments contextuels](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="7817f-p163">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="7817f-1012">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="7817f-1012">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="7817f-1013">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-1013">Requirements</span></span>

|<span data-ttu-id="7817f-1014">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-1014">Requirement</span></span>| <span data-ttu-id="7817f-1015">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-1015">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-1016">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-1016">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-1017">1.6</span><span class="sxs-lookup"><span data-stu-id="7817f-1017">-16</span></span> |
|[<span data-ttu-id="7817f-1018">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-1018">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-1019">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7817f-1019">ReadItem</span></span>|
|[<span data-ttu-id="7817f-1020">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-1020">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-1021">Lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-1021">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7817f-1022">Retourne :</span><span class="sxs-lookup"><span data-stu-id="7817f-1022">Returns:</span></span>

<span data-ttu-id="7817f-1023">Type : [Entités](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="7817f-1023">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="7817f-1024">Exemple</span><span class="sxs-lookup"><span data-stu-id="7817f-1024">Example</span></span>

<span data-ttu-id="7817f-1025">L’exemple suivant accède aux entités d’adresses dans la correspondance en surbrillance sélectionnée par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="7817f-1025">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="7817f-1026">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="7817f-1026">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="7817f-p164">Renvoie des valeurs de chaîne dans une correspondance en surbrillance, qui correspondent aux expressions régulières définies dans le fichier manifeste XML. Les correspondances en surbrillance s’appliquent aux [compléments contextuels](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="7817f-p164">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="7817f-1029">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="7817f-1029">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="7817f-p165">La méthode `getSelectedRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="7817f-p165">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="7817f-1033">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="7817f-1033">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="7817f-1034">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="7817f-1034">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="7817f-p166">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le texte. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du texte de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du texte d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du texte de l’élément.</span><span class="sxs-lookup"><span data-stu-id="7817f-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="7817f-1038">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-1038">Requirements</span></span>

|<span data-ttu-id="7817f-1039">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-1039">Requirement</span></span>| <span data-ttu-id="7817f-1040">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-1040">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-1041">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-1041">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-1042">1.6</span><span class="sxs-lookup"><span data-stu-id="7817f-1042">-16</span></span> |
|[<span data-ttu-id="7817f-1043">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-1043">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-1044">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7817f-1044">ReadItem</span></span>|
|[<span data-ttu-id="7817f-1045">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-1045">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-1046">Lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-1046">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="7817f-1047">Retourne :</span><span class="sxs-lookup"><span data-stu-id="7817f-1047">Returns:</span></span>

<span data-ttu-id="7817f-p167">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="7817f-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="7817f-1050">Exemple</span><span class="sxs-lookup"><span data-stu-id="7817f-1050">Example</span></span>

<span data-ttu-id="7817f-1051">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="7817f-1051">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="7817f-1052">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="7817f-1052">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="7817f-1053">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="7817f-1053">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="7817f-p168">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="7817f-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7817f-1057">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="7817f-1057">Parameters:</span></span>

|<span data-ttu-id="7817f-1058">Nom</span><span class="sxs-lookup"><span data-stu-id="7817f-1058">Name</span></span>| <span data-ttu-id="7817f-1059">Type</span><span class="sxs-lookup"><span data-stu-id="7817f-1059">Type</span></span>| <span data-ttu-id="7817f-1060">Attributs</span><span class="sxs-lookup"><span data-stu-id="7817f-1060">Attributes</span></span>| <span data-ttu-id="7817f-1061">Description</span><span class="sxs-lookup"><span data-stu-id="7817f-1061">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="7817f-1062">fonction</span><span class="sxs-lookup"><span data-stu-id="7817f-1062">function</span></span>||<span data-ttu-id="7817f-1063">Quand la méthode se termine, la fonction passée dans le paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7817f-1063">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="7817f-1064">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="7817f-1064">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="7817f-1065">Cet objet peut être utilisé pour obtenir, définir et supprimer les propriétés personnalisées de l’élément et sauvegarder les modifications à la propriété personnalisée est redéfinie sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="7817f-1065">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="7817f-1066">Object</span><span class="sxs-lookup"><span data-stu-id="7817f-1066">Object</span></span>| <span data-ttu-id="7817f-1067">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="7817f-1067">&lt;optional&gt;</span></span>|<span data-ttu-id="7817f-1068">Les développeurs peuvent fournir n'importe quel objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="7817f-1068">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="7817f-1069">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="7817f-1069">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7817f-1070">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-1070">Requirements</span></span>

|<span data-ttu-id="7817f-1071">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-1071">Requirement</span></span>| <span data-ttu-id="7817f-1072">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-1072">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-1073">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-1073">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-1074">1.0</span><span class="sxs-lookup"><span data-stu-id="7817f-1074">1.0</span></span>|
|[<span data-ttu-id="7817f-1075">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-1075">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-1076">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7817f-1076">ReadItem</span></span>|
|[<span data-ttu-id="7817f-1077">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-1077">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-1078">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="7817f-1078">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="7817f-1079">Exemple</span><span class="sxs-lookup"><span data-stu-id="7817f-1079">Example</span></span>

<span data-ttu-id="7817f-p171">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="7817f-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```
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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="7817f-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="7817f-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="7817f-1084">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="7817f-1084">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="7817f-p172">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément. Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session. Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session. Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer dans un formulaire en ligne qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="7817f-p172">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7817f-1089">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="7817f-1089">Parameters:</span></span>

|<span data-ttu-id="7817f-1090">Nom</span><span class="sxs-lookup"><span data-stu-id="7817f-1090">Name</span></span>| <span data-ttu-id="7817f-1091">Type</span><span class="sxs-lookup"><span data-stu-id="7817f-1091">Type</span></span>| <span data-ttu-id="7817f-1092">Attributs</span><span class="sxs-lookup"><span data-stu-id="7817f-1092">Attributes</span></span>| <span data-ttu-id="7817f-1093">Description</span><span class="sxs-lookup"><span data-stu-id="7817f-1093">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="7817f-1094">String</span><span class="sxs-lookup"><span data-stu-id="7817f-1094">String</span></span>||<span data-ttu-id="7817f-p173">Identificateur de la pièce jointe à supprimer. La longueur maximale de la chaîne est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="7817f-p173">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="7817f-1097">Object</span><span class="sxs-lookup"><span data-stu-id="7817f-1097">Object</span></span>| <span data-ttu-id="7817f-1098">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="7817f-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="7817f-1099">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="7817f-1099">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="7817f-1100">Object</span><span class="sxs-lookup"><span data-stu-id="7817f-1100">Object</span></span>| <span data-ttu-id="7817f-1101">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="7817f-1101">&lt;optional&gt;</span></span>|<span data-ttu-id="7817f-1102">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="7817f-1102">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="7817f-1103">fonction</span><span class="sxs-lookup"><span data-stu-id="7817f-1103">function</span></span>| <span data-ttu-id="7817f-1104">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="7817f-1104">&lt;optional&gt;</span></span>|<span data-ttu-id="7817f-1105">Quand la méthode se termine, la fonction passée dans le paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7817f-1105">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="7817f-1106">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="7817f-1106">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="7817f-1107">Erreurs</span><span class="sxs-lookup"><span data-stu-id="7817f-1107">Errors</span></span>

| <span data-ttu-id="7817f-1108">Code d’erreur</span><span class="sxs-lookup"><span data-stu-id="7817f-1108">Error code</span></span> | <span data-ttu-id="7817f-1109">Description</span><span class="sxs-lookup"><span data-stu-id="7817f-1109">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="7817f-1110">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="7817f-1110">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7817f-1111">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-1111">Requirements</span></span>

|<span data-ttu-id="7817f-1112">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-1112">Requirement</span></span>| <span data-ttu-id="7817f-1113">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-1113">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-1114">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-1114">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-1115">1.1</span><span class="sxs-lookup"><span data-stu-id="7817f-1115">1.1</span></span>|
|[<span data-ttu-id="7817f-1116">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-1116">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-1117">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="7817f-1117">ReadWriteItem</span></span>|
|[<span data-ttu-id="7817f-1118">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-1118">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-1119">Composition</span><span class="sxs-lookup"><span data-stu-id="7817f-1119">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="7817f-1120">Exemple</span><span class="sxs-lookup"><span data-stu-id="7817f-1120">Example</span></span>

<span data-ttu-id="7817f-1121">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="7817f-1121">The following code removes an attachment with an identifier of '0'.</span></span>

```
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="7817f-1122">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="7817f-1122">saveAsync([options], callback)</span></span>

<span data-ttu-id="7817f-1123">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="7817f-1123">Asynchronously saves an item.</span></span>

<span data-ttu-id="7817f-p174">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’identificateur de l’élément via la méthode de rappel. Dans Outlook Web App ou Outlook en mode en ligne, l’élément est enregistré sur le serveur. Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="7817f-p174">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="7817f-1127">Si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` pour utiliser avec EWS ou l’API REST, gardez à l’esprit que quand Outlook est en mode mis en cache, il peut prendre un certain temps avant que l’élément est réellement synchronisé avec le serveur.</span><span class="sxs-lookup"><span data-stu-id="7817f-1127">Note: If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the  will return an error.</span></span> <span data-ttu-id="7817f-1128">Jusqu'à ce que l’élément soit synchronisé, utiliser la propriété `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="7817f-1128">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="7817f-p176">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="7817f-p176">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="7817f-1132">Les clients suivants ont un comportement différent pour `saveAsync` pour les rendez-vous en mode composition :</span><span class="sxs-lookup"><span data-stu-id="7817f-1132">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="7817f-1133">Mac Outlook ne gère pas `saveAsync` pour une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="7817f-1133">Note: Mac Outlook does not support `saveAsync` on a meeting in compose mode. Calling  on a meeting in Mac Outlook will return an error.</span></span> <span data-ttu-id="7817f-1134">Faire appel à `saveAsync`  pour une réunion dans Outlook Mac renverra une erreur.</span><span class="sxs-lookup"><span data-stu-id="7817f-1134">Note: Mac Outlook does not support  on a meeting in compose mode. Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="7817f-1135">Outlook sur le web envoie toujours une invitation ou une mise à jour quand `saveAsync` est appelée pour un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="7817f-1135">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7817f-1136">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="7817f-1136">Parameters:</span></span>

|<span data-ttu-id="7817f-1137">Nom</span><span class="sxs-lookup"><span data-stu-id="7817f-1137">Name</span></span>| <span data-ttu-id="7817f-1138">Type</span><span class="sxs-lookup"><span data-stu-id="7817f-1138">Type</span></span>| <span data-ttu-id="7817f-1139">Attributs</span><span class="sxs-lookup"><span data-stu-id="7817f-1139">Attributes</span></span>| <span data-ttu-id="7817f-1140">Description</span><span class="sxs-lookup"><span data-stu-id="7817f-1140">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="7817f-1141">Object</span><span class="sxs-lookup"><span data-stu-id="7817f-1141">Object</span></span>| <span data-ttu-id="7817f-1142">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="7817f-1142">&lt;optional&gt;</span></span>|<span data-ttu-id="7817f-1143">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="7817f-1143">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="7817f-1144">Object</span><span class="sxs-lookup"><span data-stu-id="7817f-1144">Object</span></span>| <span data-ttu-id="7817f-1145">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="7817f-1145">&lt;optional&gt;</span></span>|<span data-ttu-id="7817f-1146">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="7817f-1146">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="7817f-1147">fonction</span><span class="sxs-lookup"><span data-stu-id="7817f-1147">function</span></span>||<span data-ttu-id="7817f-1148">Quand la méthode se termine, la fonction passée dans le paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7817f-1148">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="7817f-1149">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="7817f-1149">On success, the item identifier is provided in the `asyncResult.value` property.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7817f-1150">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-1150">Requirements</span></span>

|<span data-ttu-id="7817f-1151">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-1151">Requirement</span></span>| <span data-ttu-id="7817f-1152">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-1152">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-1153">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-1153">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-1154">1.3</span><span class="sxs-lookup"><span data-stu-id="7817f-1154">1.3</span></span>|
|[<span data-ttu-id="7817f-1155">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-1155">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-1156">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="7817f-1156">ReadWriteItem</span></span>|
|[<span data-ttu-id="7817f-1157">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-1157">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-1158">Composition</span><span class="sxs-lookup"><span data-stu-id="7817f-1158">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="7817f-1159">Exemples</span><span class="sxs-lookup"><span data-stu-id="7817f-1159">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="7817f-p178">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="7817f-p178">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="7817f-1162">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="7817f-1162">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="7817f-1163">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="7817f-1163">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="7817f-p179">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le texte ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du texte ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="7817f-p179">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="7817f-1167">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="7817f-1167">Parameters:</span></span>

|<span data-ttu-id="7817f-1168">Nom</span><span class="sxs-lookup"><span data-stu-id="7817f-1168">Name</span></span>| <span data-ttu-id="7817f-1169">Type</span><span class="sxs-lookup"><span data-stu-id="7817f-1169">Type</span></span>| <span data-ttu-id="7817f-1170">Attributs</span><span class="sxs-lookup"><span data-stu-id="7817f-1170">Attributes</span></span>| <span data-ttu-id="7817f-1171">Description</span><span class="sxs-lookup"><span data-stu-id="7817f-1171">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="7817f-1172">String</span><span class="sxs-lookup"><span data-stu-id="7817f-1172">String</span></span>||<span data-ttu-id="7817f-p180">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="7817f-p180">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="7817f-1176">Object</span><span class="sxs-lookup"><span data-stu-id="7817f-1176">Object</span></span>| <span data-ttu-id="7817f-1177">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="7817f-1177">&lt;optional&gt;</span></span>|<span data-ttu-id="7817f-1178">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="7817f-1178">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="7817f-1179">Object</span><span class="sxs-lookup"><span data-stu-id="7817f-1179">Object</span></span>| <span data-ttu-id="7817f-1180">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="7817f-1180">&lt;optional&gt;</span></span>|<span data-ttu-id="7817f-1181">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="7817f-1181">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="7817f-1182">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="7817f-1182">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="7817f-1183">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="7817f-1183">&lt;optional&gt;</span></span>|<span data-ttu-id="7817f-p181">Si `text`, le style existant est appliqué dans Outlook Web App et Outlook. Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="7817f-p181">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="7817f-p182">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook Web App et le style par défaut dans Outlook. Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="7817f-p182">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="7817f-1188">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="7817f-1188">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="7817f-1189">fonction</span><span class="sxs-lookup"><span data-stu-id="7817f-1189">function</span></span>||<span data-ttu-id="7817f-1190">Quand la méthode se termine, la fonction passée dans le paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="7817f-1190">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="7817f-1191">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-1191">Requirements</span></span>

|<span data-ttu-id="7817f-1192">Condition requise</span><span class="sxs-lookup"><span data-stu-id="7817f-1192">Requirement</span></span>| <span data-ttu-id="7817f-1193">Valeur</span><span class="sxs-lookup"><span data-stu-id="7817f-1193">Value</span></span>|
|---|---|
|[<span data-ttu-id="7817f-1194">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7817f-1194">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7817f-1195">1.2</span><span class="sxs-lookup"><span data-stu-id="7817f-1195">1.2</span></span>|
|[<span data-ttu-id="7817f-1196">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7817f-1196">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7817f-1197">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="7817f-1197">ReadWriteItem</span></span>|
|[<span data-ttu-id="7817f-1198">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7817f-1198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7817f-1199">Composition</span><span class="sxs-lookup"><span data-stu-id="7817f-1199">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="7817f-1200">Exemple</span><span class="sxs-lookup"><span data-stu-id="7817f-1200">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```