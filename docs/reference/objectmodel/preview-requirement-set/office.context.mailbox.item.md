
# <a name="item"></a><span data-ttu-id="c400f-101">élément</span><span class="sxs-lookup"><span data-stu-id="c400f-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="c400f-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="c400f-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="c400f-p101">L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype).</span><span class="sxs-lookup"><span data-stu-id="c400f-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c400f-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-105">Requirements</span></span>

|<span data-ttu-id="c400f-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-106">Requirement</span></span>|<span data-ttu-id="c400f-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-109">1.0</span><span class="sxs-lookup"><span data-stu-id="c400f-109">1.0</span></span>|
|[<span data-ttu-id="c400f-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-111">Restreinte</span><span class="sxs-lookup"><span data-stu-id="c400f-111">Restricted</span></span>|
|[<span data-ttu-id="c400f-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-113">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c400f-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="c400f-114">Members and methods</span></span>

| <span data-ttu-id="c400f-115">Membre</span><span class="sxs-lookup"><span data-stu-id="c400f-115">Member</span></span> | <span data-ttu-id="c400f-116">Type</span><span class="sxs-lookup"><span data-stu-id="c400f-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c400f-117">attachments</span><span class="sxs-lookup"><span data-stu-id="c400f-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="c400f-118">Membre</span><span class="sxs-lookup"><span data-stu-id="c400f-118">Member</span></span> |
| [<span data-ttu-id="c400f-119">bcc</span><span class="sxs-lookup"><span data-stu-id="c400f-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="c400f-120">Membre</span><span class="sxs-lookup"><span data-stu-id="c400f-120">Member</span></span> |
| [<span data-ttu-id="c400f-121">body</span><span class="sxs-lookup"><span data-stu-id="c400f-121">body</span></span>](#body-bodyjavascriptapioutlookofficebody) | <span data-ttu-id="c400f-122">Membre</span><span class="sxs-lookup"><span data-stu-id="c400f-122">Member</span></span> |
| [<span data-ttu-id="c400f-123">cc</span><span class="sxs-lookup"><span data-stu-id="c400f-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="c400f-124">Membre</span><span class="sxs-lookup"><span data-stu-id="c400f-124">Member</span></span> |
| [<span data-ttu-id="c400f-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="c400f-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="c400f-126">Membre</span><span class="sxs-lookup"><span data-stu-id="c400f-126">Member</span></span> |
| [<span data-ttu-id="c400f-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="c400f-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="c400f-128">Membre</span><span class="sxs-lookup"><span data-stu-id="c400f-128">Member</span></span> |
| [<span data-ttu-id="c400f-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="c400f-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="c400f-130">Membre</span><span class="sxs-lookup"><span data-stu-id="c400f-130">Member</span></span> |
| [<span data-ttu-id="c400f-131">end</span><span class="sxs-lookup"><span data-stu-id="c400f-131">end</span></span>](#end-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="c400f-132">Membre</span><span class="sxs-lookup"><span data-stu-id="c400f-132">Member</span></span> |
| [<span data-ttu-id="c400f-133">from</span><span class="sxs-lookup"><span data-stu-id="c400f-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | <span data-ttu-id="c400f-134">Member</span><span class="sxs-lookup"><span data-stu-id="c400f-134">Member</span></span> |
| [<span data-ttu-id="c400f-135">internetHeaders</span><span class="sxs-lookup"><span data-stu-id="c400f-135">internetHeaders</span></span>](#internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders) | <span data-ttu-id="c400f-136">Member</span><span class="sxs-lookup"><span data-stu-id="c400f-136">Member</span></span> |
| [<span data-ttu-id="c400f-137">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="c400f-137">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="c400f-138">Membre</span><span class="sxs-lookup"><span data-stu-id="c400f-138">Member</span></span> |
| [<span data-ttu-id="c400f-139">itemClass</span><span class="sxs-lookup"><span data-stu-id="c400f-139">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="c400f-140">Membre</span><span class="sxs-lookup"><span data-stu-id="c400f-140">Member</span></span> |
| [<span data-ttu-id="c400f-141">itemId</span><span class="sxs-lookup"><span data-stu-id="c400f-141">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="c400f-142">Membre</span><span class="sxs-lookup"><span data-stu-id="c400f-142">Member</span></span> |
| [<span data-ttu-id="c400f-143">itemType</span><span class="sxs-lookup"><span data-stu-id="c400f-143">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | <span data-ttu-id="c400f-144">Membre</span><span class="sxs-lookup"><span data-stu-id="c400f-144">Member</span></span> |
| [<span data-ttu-id="c400f-145">location</span><span class="sxs-lookup"><span data-stu-id="c400f-145">location</span></span>](#location-stringlocationjavascriptapioutlookofficelocation) | <span data-ttu-id="c400f-146">Membre</span><span class="sxs-lookup"><span data-stu-id="c400f-146">Member</span></span> |
| [<span data-ttu-id="c400f-147">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="c400f-147">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="c400f-148">Membre</span><span class="sxs-lookup"><span data-stu-id="c400f-148">Member</span></span> |
| [<span data-ttu-id="c400f-149">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="c400f-149">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | <span data-ttu-id="c400f-150">Membre</span><span class="sxs-lookup"><span data-stu-id="c400f-150">Member</span></span> |
| [<span data-ttu-id="c400f-151">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="c400f-151">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="c400f-152">Membre</span><span class="sxs-lookup"><span data-stu-id="c400f-152">Member</span></span> |
| [<span data-ttu-id="c400f-153">organizer</span><span class="sxs-lookup"><span data-stu-id="c400f-153">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | <span data-ttu-id="c400f-154">Member</span><span class="sxs-lookup"><span data-stu-id="c400f-154">Member</span></span> |
| [<span data-ttu-id="c400f-155">recurrence</span><span class="sxs-lookup"><span data-stu-id="c400f-155">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | <span data-ttu-id="c400f-156">Member</span><span class="sxs-lookup"><span data-stu-id="c400f-156">Member</span></span> |
| [<span data-ttu-id="c400f-157">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="c400f-157">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="c400f-158">Member</span><span class="sxs-lookup"><span data-stu-id="c400f-158">Member</span></span> |
| [<span data-ttu-id="c400f-159">sender</span><span class="sxs-lookup"><span data-stu-id="c400f-159">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | <span data-ttu-id="c400f-160">Member</span><span class="sxs-lookup"><span data-stu-id="c400f-160">Member</span></span> |
| [<span data-ttu-id="c400f-161">seriesId</span><span class="sxs-lookup"><span data-stu-id="c400f-161">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="c400f-162">Member</span><span class="sxs-lookup"><span data-stu-id="c400f-162">Member</span></span> |
| [<span data-ttu-id="c400f-163">start</span><span class="sxs-lookup"><span data-stu-id="c400f-163">start</span></span>](#start-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="c400f-164">Membre</span><span class="sxs-lookup"><span data-stu-id="c400f-164">Member</span></span> |
| [<span data-ttu-id="c400f-165">subject</span><span class="sxs-lookup"><span data-stu-id="c400f-165">subject</span></span>](#subject-stringsubjectjavascriptapioutlookofficesubject) | <span data-ttu-id="c400f-166">Membre</span><span class="sxs-lookup"><span data-stu-id="c400f-166">Member</span></span> |
| [<span data-ttu-id="c400f-167">to</span><span class="sxs-lookup"><span data-stu-id="c400f-167">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="c400f-168">Membre</span><span class="sxs-lookup"><span data-stu-id="c400f-168">Member</span></span> |
| [<span data-ttu-id="c400f-169">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="c400f-169">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="c400f-170">Méthode</span><span class="sxs-lookup"><span data-stu-id="c400f-170">Method</span></span> |
| [<span data-ttu-id="c400f-171">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="c400f-171">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="c400f-172">Méthode</span><span class="sxs-lookup"><span data-stu-id="c400f-172">Method</span></span> |
| [<span data-ttu-id="c400f-173">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="c400f-173">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="c400f-174">Méthode</span><span class="sxs-lookup"><span data-stu-id="c400f-174">Method</span></span> |
| [<span data-ttu-id="c400f-175">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="c400f-175">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="c400f-176">Méthode</span><span class="sxs-lookup"><span data-stu-id="c400f-176">Method</span></span> |
| [<span data-ttu-id="c400f-177">close</span><span class="sxs-lookup"><span data-stu-id="c400f-177">close</span></span>](#close) | <span data-ttu-id="c400f-178">Méthode</span><span class="sxs-lookup"><span data-stu-id="c400f-178">Method</span></span> |
| [<span data-ttu-id="c400f-179">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="c400f-179">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="c400f-180">Méthode</span><span class="sxs-lookup"><span data-stu-id="c400f-180">Method</span></span> |
| [<span data-ttu-id="c400f-181">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="c400f-181">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="c400f-182">Méthode</span><span class="sxs-lookup"><span data-stu-id="c400f-182">Method</span></span> |
| [<span data-ttu-id="c400f-183">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="c400f-183">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent) | <span data-ttu-id="c400f-184">Méthode</span><span class="sxs-lookup"><span data-stu-id="c400f-184">Method</span></span> |
| [<span data-ttu-id="c400f-185">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="c400f-185">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="c400f-186">Méthode</span><span class="sxs-lookup"><span data-stu-id="c400f-186">Method</span></span> |
| [<span data-ttu-id="c400f-187">getEntities</span><span class="sxs-lookup"><span data-stu-id="c400f-187">getEntities</span></span>](#getentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="c400f-188">Méthode</span><span class="sxs-lookup"><span data-stu-id="c400f-188">Method</span></span> |
| [<span data-ttu-id="c400f-189">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="c400f-189">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="c400f-190">Méthode</span><span class="sxs-lookup"><span data-stu-id="c400f-190">Method</span></span> |
| [<span data-ttu-id="c400f-191">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="c400f-191">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="c400f-192">Méthode</span><span class="sxs-lookup"><span data-stu-id="c400f-192">Method</span></span> |
| [<span data-ttu-id="c400f-193">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="c400f-193">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="c400f-194">Méthode</span><span class="sxs-lookup"><span data-stu-id="c400f-194">Method</span></span> |
| [<span data-ttu-id="c400f-195">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="c400f-195">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="c400f-196">Méthode</span><span class="sxs-lookup"><span data-stu-id="c400f-196">Method</span></span> |
| [<span data-ttu-id="c400f-197">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="c400f-197">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="c400f-198">Méthode</span><span class="sxs-lookup"><span data-stu-id="c400f-198">Method</span></span> |
| [<span data-ttu-id="c400f-199">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="c400f-199">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="c400f-200">Méthode</span><span class="sxs-lookup"><span data-stu-id="c400f-200">Method</span></span> |
| [<span data-ttu-id="c400f-201">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="c400f-201">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="c400f-202">Méthode</span><span class="sxs-lookup"><span data-stu-id="c400f-202">Method</span></span> |
| [<span data-ttu-id="c400f-203">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="c400f-203">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="c400f-204">Méthode</span><span class="sxs-lookup"><span data-stu-id="c400f-204">Method</span></span> |
| [<span data-ttu-id="c400f-205">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="c400f-205">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="c400f-206">Méthode</span><span class="sxs-lookup"><span data-stu-id="c400f-206">Method</span></span> |
| [<span data-ttu-id="c400f-207">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="c400f-207">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="c400f-208">Méthode</span><span class="sxs-lookup"><span data-stu-id="c400f-208">Method</span></span> |
| [<span data-ttu-id="c400f-209">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="c400f-209">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="c400f-210">Méthode</span><span class="sxs-lookup"><span data-stu-id="c400f-210">Method</span></span> |
| [<span data-ttu-id="c400f-211">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="c400f-211">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="c400f-212">Méthode</span><span class="sxs-lookup"><span data-stu-id="c400f-212">Method</span></span> |
| [<span data-ttu-id="c400f-213">saveAsync</span><span class="sxs-lookup"><span data-stu-id="c400f-213">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="c400f-214">Méthode</span><span class="sxs-lookup"><span data-stu-id="c400f-214">Method</span></span> |
| [<span data-ttu-id="c400f-215">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="c400f-215">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="c400f-216">Méthode</span><span class="sxs-lookup"><span data-stu-id="c400f-216">Method</span></span> |

### <a name="example"></a><span data-ttu-id="c400f-217">Exemple</span><span class="sxs-lookup"><span data-stu-id="c400f-217">Example</span></span>

<span data-ttu-id="c400f-218">L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="c400f-218">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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
}
```

### <a name="members"></a><span data-ttu-id="c400f-219">Membres</span><span class="sxs-lookup"><span data-stu-id="c400f-219">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="c400f-220">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="c400f-220">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="c400f-221">Permet d’obtenir les pièces jointes de l’élément sous forme de tableau.</span><span class="sxs-lookup"><span data-stu-id="c400f-221">Gets the item's attachments as an array.</span></span> <span data-ttu-id="c400f-222">Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="c400f-222">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c400f-223">Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés.</span><span class="sxs-lookup"><span data-stu-id="c400f-223">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="c400f-224">Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span><span class="sxs-lookup"><span data-stu-id="c400f-224">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="c400f-225">Type :</span><span class="sxs-lookup"><span data-stu-id="c400f-225">Type:</span></span>

*   <span data-ttu-id="c400f-226">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="c400f-226">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="c400f-227">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-227">Requirements</span></span>

|<span data-ttu-id="c400f-228">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-228">Requirement</span></span>|<span data-ttu-id="c400f-229">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-229">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-230">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-230">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-231">1.0</span><span class="sxs-lookup"><span data-stu-id="c400f-231">1.0</span></span>|
|[<span data-ttu-id="c400f-232">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-232">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-233">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-233">ReadItem</span></span>|
|[<span data-ttu-id="c400f-234">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-234">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-235">Lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-235">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c400f-236">Exemple</span><span class="sxs-lookup"><span data-stu-id="c400f-236">Example</span></span>

<span data-ttu-id="c400f-237">Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="c400f-237">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```javascript
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

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="c400f-238">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c400f-238">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="c400f-239">Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message.</span><span class="sxs-lookup"><span data-stu-id="c400f-239">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="c400f-240">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="c400f-240">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c400f-241">Type :</span><span class="sxs-lookup"><span data-stu-id="c400f-241">Type:</span></span>

*   [<span data-ttu-id="c400f-242">Destinataires</span><span class="sxs-lookup"><span data-stu-id="c400f-242">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="c400f-243">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-243">Requirements</span></span>

|<span data-ttu-id="c400f-244">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-244">Requirement</span></span>|<span data-ttu-id="c400f-245">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-245">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-246">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-246">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-247">1.1</span><span class="sxs-lookup"><span data-stu-id="c400f-247">1.1</span></span>|
|[<span data-ttu-id="c400f-248">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-248">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-249">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-249">ReadItem</span></span>|
|[<span data-ttu-id="c400f-250">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-250">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-251">Composition</span><span class="sxs-lookup"><span data-stu-id="c400f-251">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c400f-252">Exemple</span><span class="sxs-lookup"><span data-stu-id="c400f-252">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="c400f-253">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="c400f-253">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="c400f-254">Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.</span><span class="sxs-lookup"><span data-stu-id="c400f-254">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="c400f-255">Type :</span><span class="sxs-lookup"><span data-stu-id="c400f-255">Type:</span></span>

*   [<span data-ttu-id="c400f-256">Corps</span><span class="sxs-lookup"><span data-stu-id="c400f-256">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="c400f-257">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-257">Requirements</span></span>

|<span data-ttu-id="c400f-258">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-258">Requirement</span></span>|<span data-ttu-id="c400f-259">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-259">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-260">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-260">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-261">1.1</span><span class="sxs-lookup"><span data-stu-id="c400f-261">1.1</span></span>|
|[<span data-ttu-id="c400f-262">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-262">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-263">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-263">ReadItem</span></span>|
|[<span data-ttu-id="c400f-264">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-264">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-265">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-265">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="c400f-266">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c400f-266">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="c400f-267">Permet d’accéder aux destinataires en copie carbone (Cc) d’un message.</span><span class="sxs-lookup"><span data-stu-id="c400f-267">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="c400f-268">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="c400f-268">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c400f-269">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-269">Read mode</span></span>

<span data-ttu-id="c400f-p106">La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="c400f-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c400f-272">Mode composition</span><span class="sxs-lookup"><span data-stu-id="c400f-272">Compose mode</span></span>

<span data-ttu-id="c400f-273">La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message.</span><span class="sxs-lookup"><span data-stu-id="c400f-273">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="c400f-274">Type :</span><span class="sxs-lookup"><span data-stu-id="c400f-274">Type:</span></span>

*   <span data-ttu-id="c400f-275">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c400f-275">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c400f-276">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-276">Requirements</span></span>

|<span data-ttu-id="c400f-277">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-277">Requirement</span></span>|<span data-ttu-id="c400f-278">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-278">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-279">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-279">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-280">1.0</span><span class="sxs-lookup"><span data-stu-id="c400f-280">1.0</span></span>|
|[<span data-ttu-id="c400f-281">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-281">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-282">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-282">ReadItem</span></span>|
|[<span data-ttu-id="c400f-283">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-283">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-284">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-284">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c400f-285">Exemple</span><span class="sxs-lookup"><span data-stu-id="c400f-285">Example</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="c400f-286">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="c400f-286">(nullable) conversationId :String</span></span>

<span data-ttu-id="c400f-287">Obtient l’identificateur de la conversation qui contient un message particulier.</span><span class="sxs-lookup"><span data-stu-id="c400f-287">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="c400f-p107">Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.</span><span class="sxs-lookup"><span data-stu-id="c400f-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="c400f-p108">Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.</span><span class="sxs-lookup"><span data-stu-id="c400f-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="c400f-292">Type :</span><span class="sxs-lookup"><span data-stu-id="c400f-292">Type:</span></span>

*   <span data-ttu-id="c400f-293">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c400f-293">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c400f-294">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-294">Requirements</span></span>

|<span data-ttu-id="c400f-295">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-295">Requirement</span></span>|<span data-ttu-id="c400f-296">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-296">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-297">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-297">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-298">1.0</span><span class="sxs-lookup"><span data-stu-id="c400f-298">1.0</span></span>|
|[<span data-ttu-id="c400f-299">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-299">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-300">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-300">ReadItem</span></span>|
|[<span data-ttu-id="c400f-301">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-301">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-302">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-302">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="c400f-303">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="c400f-303">dateTimeCreated :Date</span></span>

<span data-ttu-id="c400f-p109">Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="c400f-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c400f-306">Type :</span><span class="sxs-lookup"><span data-stu-id="c400f-306">Type:</span></span>

*   <span data-ttu-id="c400f-307">Date</span><span class="sxs-lookup"><span data-stu-id="c400f-307">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="c400f-308">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-308">Requirements</span></span>

|<span data-ttu-id="c400f-309">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-309">Requirement</span></span>|<span data-ttu-id="c400f-310">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-311">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-312">1.0</span><span class="sxs-lookup"><span data-stu-id="c400f-312">1.0</span></span>|
|[<span data-ttu-id="c400f-313">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-313">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-314">ReadItem</span></span>|
|[<span data-ttu-id="c400f-315">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-315">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-316">Lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-316">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c400f-317">Exemple</span><span class="sxs-lookup"><span data-stu-id="c400f-317">Example</span></span>

```javascript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="c400f-318">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="c400f-318">dateTimeModified :Date</span></span>

<span data-ttu-id="c400f-p110">Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="c400f-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c400f-321">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="c400f-321">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="c400f-322">Type :</span><span class="sxs-lookup"><span data-stu-id="c400f-322">Type:</span></span>

*   <span data-ttu-id="c400f-323">Date</span><span class="sxs-lookup"><span data-stu-id="c400f-323">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="c400f-324">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-324">Requirements</span></span>

|<span data-ttu-id="c400f-325">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-325">Requirement</span></span>|<span data-ttu-id="c400f-326">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-326">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-327">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-327">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-328">1.0</span><span class="sxs-lookup"><span data-stu-id="c400f-328">1.0</span></span>|
|[<span data-ttu-id="c400f-329">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-329">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-330">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-330">ReadItem</span></span>|
|[<span data-ttu-id="c400f-331">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-331">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-332">Lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-332">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c400f-333">Exemple</span><span class="sxs-lookup"><span data-stu-id="c400f-333">Example</span></span>

```javascript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="c400f-334">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="c400f-334">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="c400f-335">Obtient ou définit la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="c400f-335">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="c400f-p111">La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="c400f-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c400f-338">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-338">Read mode</span></span>

<span data-ttu-id="c400f-339">La propriété `end` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="c400f-339">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c400f-340">Mode composition</span><span class="sxs-lookup"><span data-stu-id="c400f-340">Compose mode</span></span>

<span data-ttu-id="c400f-341">La propriété `end` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="c400f-341">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="c400f-342">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="c400f-342">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="c400f-343">Type :</span><span class="sxs-lookup"><span data-stu-id="c400f-343">Type:</span></span>

*   <span data-ttu-id="c400f-344">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="c400f-344">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c400f-345">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-345">Requirements</span></span>

|<span data-ttu-id="c400f-346">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-346">Requirement</span></span>|<span data-ttu-id="c400f-347">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-347">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-348">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-348">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-349">1.0</span><span class="sxs-lookup"><span data-stu-id="c400f-349">1.0</span></span>|
|[<span data-ttu-id="c400f-350">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-350">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-351">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-351">ReadItem</span></span>|
|[<span data-ttu-id="c400f-352">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-352">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-353">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-353">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c400f-354">Exemple</span><span class="sxs-lookup"><span data-stu-id="c400f-354">Example</span></span>

<span data-ttu-id="c400f-355">L’exemple suivant définit l’heure de fin d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="c400f-355">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```javascript
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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="c400f-356">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="c400f-356">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="c400f-357">Permet d’obtenir l’adresse de messagerie de l’expéditeur d’un message.</span><span class="sxs-lookup"><span data-stu-id="c400f-357">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="c400f-p112">Les propriétés `from` et [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="c400f-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="c400f-360">la propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="c400f-360">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c400f-361">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-361">Read mode</span></span>

<span data-ttu-id="c400f-362">La propriété `from` renvoie un objet `EmailAddressDetails`.</span><span class="sxs-lookup"><span data-stu-id="c400f-362">The `from` property returns a `EmailAddressDetails` object.</span></span>

```javascript
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="c400f-363">Mode composition</span><span class="sxs-lookup"><span data-stu-id="c400f-363">Compose mode</span></span>

<span data-ttu-id="c400f-364">La propriété `from` renvoie un objet `From` qui fournit une méthode pour obtenir la valeur from.</span><span class="sxs-lookup"><span data-stu-id="c400f-364">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c400f-365">Type :</span><span class="sxs-lookup"><span data-stu-id="c400f-365">Type:</span></span>

*   <span data-ttu-id="c400f-366">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="c400f-366">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c400f-367">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-367">Requirements</span></span>

|<span data-ttu-id="c400f-368">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-368">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="c400f-369">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-369">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-370">1.0</span><span class="sxs-lookup"><span data-stu-id="c400f-370">1.0</span></span>|<span data-ttu-id="c400f-371">1.7</span><span class="sxs-lookup"><span data-stu-id="c400f-371">-17</span></span>|
|[<span data-ttu-id="c400f-372">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-372">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-373">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-373">ReadItem</span></span>|<span data-ttu-id="c400f-374">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c400f-374">ReadWriteItem</span></span>|
|[<span data-ttu-id="c400f-375">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-375">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-376">Lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-376">Read</span></span>|<span data-ttu-id="c400f-377">Composition</span><span class="sxs-lookup"><span data-stu-id="c400f-377">Compose</span></span>|

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="c400f-378">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="c400f-378">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="c400f-379">Permet d’obtenir ou de définir les en-têtes Internet d’un message.</span><span class="sxs-lookup"><span data-stu-id="c400f-379">Gets or sets the internet headers of a message.</span></span>

##### <a name="type"></a><span data-ttu-id="c400f-380">Type :</span><span class="sxs-lookup"><span data-stu-id="c400f-380">Type:</span></span>

*   [<span data-ttu-id="c400f-381">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="c400f-381">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="c400f-382">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-382">Requirements</span></span>

|<span data-ttu-id="c400f-383">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-383">Requirement</span></span>|<span data-ttu-id="c400f-384">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-384">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-385">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-385">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-386">Aperçu</span><span class="sxs-lookup"><span data-stu-id="c400f-386">Preview</span></span>|
|[<span data-ttu-id="c400f-387">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-387">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-388">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-388">ReadItem</span></span>|
|[<span data-ttu-id="c400f-389">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-389">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-390">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-390">Compose or read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="c400f-391">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="c400f-391">internetMessageId :String</span></span>

<span data-ttu-id="c400f-p113">Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="c400f-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c400f-394">Type :</span><span class="sxs-lookup"><span data-stu-id="c400f-394">Type:</span></span>

*   <span data-ttu-id="c400f-395">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c400f-395">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c400f-396">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-396">Requirements</span></span>

|<span data-ttu-id="c400f-397">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-397">Requirement</span></span>|<span data-ttu-id="c400f-398">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-398">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-399">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-399">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-400">1.0</span><span class="sxs-lookup"><span data-stu-id="c400f-400">1.0</span></span>|
|[<span data-ttu-id="c400f-401">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-401">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-402">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-402">ReadItem</span></span>|
|[<span data-ttu-id="c400f-403">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-403">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-404">Lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-404">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c400f-405">Exemple</span><span class="sxs-lookup"><span data-stu-id="c400f-405">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="c400f-406">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="c400f-406">itemClass :String</span></span>

<span data-ttu-id="c400f-p114">Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="c400f-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="c400f-p115">La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="c400f-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="c400f-411">Type</span><span class="sxs-lookup"><span data-stu-id="c400f-411">Type</span></span>|<span data-ttu-id="c400f-412">Description</span><span class="sxs-lookup"><span data-stu-id="c400f-412">Description</span></span>|<span data-ttu-id="c400f-413">Classe de l’élément</span><span class="sxs-lookup"><span data-stu-id="c400f-413">item class</span></span>|
|---|---|---|
|<span data-ttu-id="c400f-414">Éléments de rendez-vous</span><span class="sxs-lookup"><span data-stu-id="c400f-414">Appointment items</span></span>|<span data-ttu-id="c400f-415">Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurence`.</span><span class="sxs-lookup"><span data-stu-id="c400f-415">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="c400f-416">Éléments de message</span><span class="sxs-lookup"><span data-stu-id="c400f-416">Message items</span></span>|<span data-ttu-id="c400f-417">Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.</span><span class="sxs-lookup"><span data-stu-id="c400f-417">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="c400f-418">Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.</span><span class="sxs-lookup"><span data-stu-id="c400f-418">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="c400f-419">Type :</span><span class="sxs-lookup"><span data-stu-id="c400f-419">Type:</span></span>

*   <span data-ttu-id="c400f-420">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c400f-420">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c400f-421">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-421">Requirements</span></span>

|<span data-ttu-id="c400f-422">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-422">Requirement</span></span>|<span data-ttu-id="c400f-423">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-424">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-425">1.0</span><span class="sxs-lookup"><span data-stu-id="c400f-425">1.0</span></span>|
|[<span data-ttu-id="c400f-426">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-427">ReadItem</span></span>|
|[<span data-ttu-id="c400f-428">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-429">Lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-429">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c400f-430">Exemple</span><span class="sxs-lookup"><span data-stu-id="c400f-430">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="c400f-431">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="c400f-431">(nullable) itemId :String</span></span>

<span data-ttu-id="c400f-p116">Permet d’obtenir l’identificateur de l’élément des services web Exchange pour l’élément actif. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="c400f-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c400f-434">L’identificateur renvoyé par la propriété `itemId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="c400f-434">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="c400f-435">La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="c400f-435">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="c400f-436">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="c400f-436">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="c400f-437">Pour plus d’informations, consultez la rubrique [Utilisation des API REST Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span><span class="sxs-lookup"><span data-stu-id="c400f-437">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="c400f-p118">La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="c400f-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="c400f-440">Type :</span><span class="sxs-lookup"><span data-stu-id="c400f-440">Type:</span></span>

*   <span data-ttu-id="c400f-441">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c400f-441">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c400f-442">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-442">Requirements</span></span>

|<span data-ttu-id="c400f-443">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-443">Requirement</span></span>|<span data-ttu-id="c400f-444">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-444">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-445">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-445">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-446">1.0</span><span class="sxs-lookup"><span data-stu-id="c400f-446">1.0</span></span>|
|[<span data-ttu-id="c400f-447">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-447">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-448">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-448">ReadItem</span></span>|
|[<span data-ttu-id="c400f-449">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-449">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-450">Lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-450">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c400f-451">Exemple</span><span class="sxs-lookup"><span data-stu-id="c400f-451">Example</span></span>

<span data-ttu-id="c400f-p119">Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.</span><span class="sxs-lookup"><span data-stu-id="c400f-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="c400f-454">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="c400f-454">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="c400f-455">Obtient le type d’élément représenté par une instance.</span><span class="sxs-lookup"><span data-stu-id="c400f-455">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="c400f-456">La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="c400f-456">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="c400f-457">Type :</span><span class="sxs-lookup"><span data-stu-id="c400f-457">Type:</span></span>

*   [<span data-ttu-id="c400f-458">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="c400f-458">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="c400f-459">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-459">Requirements</span></span>

|<span data-ttu-id="c400f-460">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-460">Requirement</span></span>|<span data-ttu-id="c400f-461">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-461">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-462">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-462">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-463">1.0</span><span class="sxs-lookup"><span data-stu-id="c400f-463">1.0</span></span>|
|[<span data-ttu-id="c400f-464">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-464">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-465">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-465">ReadItem</span></span>|
|[<span data-ttu-id="c400f-466">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-466">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-467">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-467">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c400f-468">Exemple</span><span class="sxs-lookup"><span data-stu-id="c400f-468">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="c400f-469">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="c400f-469">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="c400f-470">Obtient ou définit le lieu d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="c400f-470">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c400f-471">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-471">Read mode</span></span>

<span data-ttu-id="c400f-472">La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="c400f-472">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c400f-473">Mode composition</span><span class="sxs-lookup"><span data-stu-id="c400f-473">Compose mode</span></span>

<span data-ttu-id="c400f-474">La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="c400f-474">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="c400f-475">Type :</span><span class="sxs-lookup"><span data-stu-id="c400f-475">Type:</span></span>

*   <span data-ttu-id="c400f-476">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="c400f-476">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c400f-477">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-477">Requirements</span></span>

|<span data-ttu-id="c400f-478">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-478">Requirement</span></span>|<span data-ttu-id="c400f-479">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-479">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-480">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-480">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-481">1.0</span><span class="sxs-lookup"><span data-stu-id="c400f-481">1.0</span></span>|
|[<span data-ttu-id="c400f-482">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-482">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-483">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-483">ReadItem</span></span>|
|[<span data-ttu-id="c400f-484">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-484">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-485">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-485">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c400f-486">Exemple</span><span class="sxs-lookup"><span data-stu-id="c400f-486">Example</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="c400f-487">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="c400f-487">normalizedSubject :String</span></span>

<span data-ttu-id="c400f-p120">Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="c400f-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="c400f-p121">La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject).</span><span class="sxs-lookup"><span data-stu-id="c400f-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="c400f-492">Type :</span><span class="sxs-lookup"><span data-stu-id="c400f-492">Type:</span></span>

*   <span data-ttu-id="c400f-493">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c400f-493">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c400f-494">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-494">Requirements</span></span>

|<span data-ttu-id="c400f-495">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-495">Requirement</span></span>|<span data-ttu-id="c400f-496">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-496">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-497">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-497">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-498">1.0</span><span class="sxs-lookup"><span data-stu-id="c400f-498">1.0</span></span>|
|[<span data-ttu-id="c400f-499">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-499">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-500">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-500">ReadItem</span></span>|
|[<span data-ttu-id="c400f-501">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-501">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-502">Lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-502">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c400f-503">Exemple</span><span class="sxs-lookup"><span data-stu-id="c400f-503">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="c400f-504">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="c400f-504">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="c400f-505">Obtient les messages de notification pour un élément.</span><span class="sxs-lookup"><span data-stu-id="c400f-505">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="c400f-506">Type :</span><span class="sxs-lookup"><span data-stu-id="c400f-506">Type:</span></span>

*   [<span data-ttu-id="c400f-507">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="c400f-507">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="c400f-508">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-508">Requirements</span></span>

|<span data-ttu-id="c400f-509">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-509">Requirement</span></span>|<span data-ttu-id="c400f-510">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-510">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-511">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-511">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-512">1.3</span><span class="sxs-lookup"><span data-stu-id="c400f-512">1.3</span></span>|
|[<span data-ttu-id="c400f-513">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-513">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-514">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-514">ReadItem</span></span>|
|[<span data-ttu-id="c400f-515">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-515">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-516">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-516">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="c400f-517">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c400f-517">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="c400f-518">Permet d’accéder aux participants facultatifs d’un événement.</span><span class="sxs-lookup"><span data-stu-id="c400f-518">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="c400f-519">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="c400f-519">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c400f-520">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-520">Read mode</span></span>

<span data-ttu-id="c400f-521">La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion.</span><span class="sxs-lookup"><span data-stu-id="c400f-521">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c400f-522">Mode composition</span><span class="sxs-lookup"><span data-stu-id="c400f-522">Compose mode</span></span>

<span data-ttu-id="c400f-523">La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion.</span><span class="sxs-lookup"><span data-stu-id="c400f-523">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="c400f-524">Type :</span><span class="sxs-lookup"><span data-stu-id="c400f-524">Type:</span></span>

*   <span data-ttu-id="c400f-525">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c400f-525">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c400f-526">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-526">Requirements</span></span>

|<span data-ttu-id="c400f-527">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-527">Requirement</span></span>|<span data-ttu-id="c400f-528">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-528">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-529">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-529">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-530">1.0</span><span class="sxs-lookup"><span data-stu-id="c400f-530">1.0</span></span>|
|[<span data-ttu-id="c400f-531">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-531">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-532">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-532">ReadItem</span></span>|
|[<span data-ttu-id="c400f-533">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-533">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-534">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-534">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c400f-535">Exemple</span><span class="sxs-lookup"><span data-stu-id="c400f-535">Example</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="c400f-536">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="c400f-536">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="c400f-537">Permet d’obtenir l’adresse de messagerie de l’organisateur d’une réunion spécifiée.</span><span class="sxs-lookup"><span data-stu-id="c400f-537">Gets the email address of the meeting organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c400f-538">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-538">Read mode</span></span>

<span data-ttu-id="c400f-539">La propriété `organizer` renvoie un objet [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) qui représente l’organisateur de la réunion.</span><span class="sxs-lookup"><span data-stu-id="c400f-539">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c400f-540">Mode composition</span><span class="sxs-lookup"><span data-stu-id="c400f-540">Compose mode</span></span>

<span data-ttu-id="c400f-541">La propriété `organizer` renvoie un objet [Organizer](/javascript/api/outlook/office.organizer) qui fournit une méthode pour obtenir la valeur organizer.</span><span class="sxs-lookup"><span data-stu-id="c400f-541">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="c400f-542">Type :</span><span class="sxs-lookup"><span data-stu-id="c400f-542">Type:</span></span>

*   <span data-ttu-id="c400f-543">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="c400f-543">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c400f-544">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-544">Requirements</span></span>

|<span data-ttu-id="c400f-545">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-545">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="c400f-546">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-546">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-547">1.0</span><span class="sxs-lookup"><span data-stu-id="c400f-547">1.0</span></span>|<span data-ttu-id="c400f-548">1.7</span><span class="sxs-lookup"><span data-stu-id="c400f-548">-17</span></span>|
|[<span data-ttu-id="c400f-549">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-549">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-550">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-550">ReadItem</span></span>|<span data-ttu-id="c400f-551">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c400f-551">ReadWriteItem</span></span>|
|[<span data-ttu-id="c400f-552">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-552">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-553">Lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-553">Read</span></span>|<span data-ttu-id="c400f-554">Composition</span><span class="sxs-lookup"><span data-stu-id="c400f-554">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c400f-555">Exemple</span><span class="sxs-lookup"><span data-stu-id="c400f-555">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="c400f-556">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="c400f-556">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="c400f-557">Permet d’obtenir ou définit la périodicité d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="c400f-557">Gets or sets the location of an appointment.</span></span> <span data-ttu-id="c400f-558">Permet d’obtenir la périodicité d’une demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="c400f-558">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="c400f-559">Modes lecture et composition pour les éléments de rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="c400f-559">Read and compose modes for appointment items.</span></span> <span data-ttu-id="c400f-560">Mode lecture pour les éléments de demande de réunion.</span><span class="sxs-lookup"><span data-stu-id="c400f-560">Read mode for meeting request items.</span></span>

<span data-ttu-id="c400f-561">La propriété `recurrence` renvoie un objet [périodicité](/javascript/api/outlook/office.recurrence) pour des demandes de réunions ou de rendez-vous périodiques si un élément est une série ou une instance dans une série.</span><span class="sxs-lookup"><span data-stu-id="c400f-561">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="c400f-562">La valeur `null` est renvoyée pour les rendez-vous uniques et les demandes de réunion de rendez-vous uniques.</span><span class="sxs-lookup"><span data-stu-id="c400f-562">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="c400f-563">La valeur `undefined` est renvoyée pour les messages qui ne sont pas des demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="c400f-563">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="c400f-564">Remarque : les demandes de réunion ont une valeur `itemClass` d’IPM. Schedule.Meeting.Request.</span><span class="sxs-lookup"><span data-stu-id="c400f-564">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="c400f-565">Remarque : si l’objet de périodicité est `null`, cela indique que l’objet est un rendez-vous unique ou une demande de réunion de rendez-vous unique, et NON une partie d’une série.</span><span class="sxs-lookup"><span data-stu-id="c400f-565">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="c400f-566">Type :</span><span class="sxs-lookup"><span data-stu-id="c400f-566">Type:</span></span>

* [<span data-ttu-id="c400f-567">Recurrence</span><span class="sxs-lookup"><span data-stu-id="c400f-567">recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="c400f-568">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-568">Requirement</span></span>|<span data-ttu-id="c400f-569">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-569">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-570">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-570">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-571">1.7</span><span class="sxs-lookup"><span data-stu-id="c400f-571">-17</span></span>|
|[<span data-ttu-id="c400f-572">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-572">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-573">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-573">ReadItem</span></span>|
|[<span data-ttu-id="c400f-574">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-574">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-575">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-575">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="c400f-576">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c400f-576">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="c400f-577">Permet d’accéder aux participants requis à un événement.</span><span class="sxs-lookup"><span data-stu-id="c400f-577">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="c400f-578">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="c400f-578">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c400f-579">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-579">Read mode</span></span>

<span data-ttu-id="c400f-580">La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion.</span><span class="sxs-lookup"><span data-stu-id="c400f-580">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c400f-581">Mode composition</span><span class="sxs-lookup"><span data-stu-id="c400f-581">Compose mode</span></span>

<span data-ttu-id="c400f-582">La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion.</span><span class="sxs-lookup"><span data-stu-id="c400f-582">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="c400f-583">Type :</span><span class="sxs-lookup"><span data-stu-id="c400f-583">Type:</span></span>

*   <span data-ttu-id="c400f-584">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c400f-584">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c400f-585">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-585">Requirements</span></span>

|<span data-ttu-id="c400f-586">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-586">Requirement</span></span>|<span data-ttu-id="c400f-587">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-587">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-588">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-588">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-589">1.0</span><span class="sxs-lookup"><span data-stu-id="c400f-589">1.0</span></span>|
|[<span data-ttu-id="c400f-590">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-590">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-591">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-591">ReadItem</span></span>|
|[<span data-ttu-id="c400f-592">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-592">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-593">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-593">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c400f-594">Exemple</span><span class="sxs-lookup"><span data-stu-id="c400f-594">Example</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="c400f-595">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="c400f-595">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="c400f-p126">Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="c400f-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="c400f-p127">Les propriétés [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.</span><span class="sxs-lookup"><span data-stu-id="c400f-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="c400f-600">la propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.</span><span class="sxs-lookup"><span data-stu-id="c400f-600">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="c400f-601">Type :</span><span class="sxs-lookup"><span data-stu-id="c400f-601">Type:</span></span>

*   [<span data-ttu-id="c400f-602">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="c400f-602">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="c400f-603">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-603">Requirements</span></span>

|<span data-ttu-id="c400f-604">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-604">Requirement</span></span>|<span data-ttu-id="c400f-605">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-605">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-606">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-607">1.0</span><span class="sxs-lookup"><span data-stu-id="c400f-607">1.0</span></span>|
|[<span data-ttu-id="c400f-608">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-608">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-609">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-609">ReadItem</span></span>|
|[<span data-ttu-id="c400f-610">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-610">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-611">Lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-611">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c400f-612">Exemple</span><span class="sxs-lookup"><span data-stu-id="c400f-612">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="c400f-613">(nullable) seriesId :String</span><span class="sxs-lookup"><span data-stu-id="c400f-613">(nullable) seriesId :String</span></span>

<span data-ttu-id="c400f-614">Permet d’obtenir l’ID de la série à laquelle une instance appartient.</span><span class="sxs-lookup"><span data-stu-id="c400f-614">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="c400f-615">Dans OWA et Outlook, `seriesId` renvoie l’identificateur de services web Exchange (EWS) de l’élément (series) parent auquel cet élément appartient.</span><span class="sxs-lookup"><span data-stu-id="c400f-615">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="c400f-616">Dans iOS et Android, `seriesId` renvoie l’ID REST de l’élément parent.</span><span class="sxs-lookup"><span data-stu-id="c400f-616">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="c400f-617">L’identificateur renvoyé par la propriété `seriesId` est identique à celui de l’élément des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="c400f-617">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="c400f-618">La propriété `seriesId` n’est pas identique aux ID Outlook utilisés par l’API REST Outlook.</span><span class="sxs-lookup"><span data-stu-id="c400f-618">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="c400f-619">Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span><span class="sxs-lookup"><span data-stu-id="c400f-619">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="c400f-620">Pour plus d’informations, consultez la rubrique [Utilisation des API REST Outlook à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span><span class="sxs-lookup"><span data-stu-id="c400f-620">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="c400f-621">La propriété `seriesId` renvoie `null` pour les éléments qui n’ont pas d’élément parent, tels que des rendez-vous uniques, des éléments de séries ou des demandes de réunion, et renvoie `undefined` pour tous les autres éléments qui ne sont pas des demandes de réunion.</span><span class="sxs-lookup"><span data-stu-id="c400f-621">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="c400f-622">Type :</span><span class="sxs-lookup"><span data-stu-id="c400f-622">Type:</span></span>

* <span data-ttu-id="c400f-623">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c400f-623">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c400f-624">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-624">Requirements</span></span>

|<span data-ttu-id="c400f-625">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-625">Requirement</span></span>|<span data-ttu-id="c400f-626">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-626">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-627">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-627">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-628">1.7</span><span class="sxs-lookup"><span data-stu-id="c400f-628">-17</span></span>|
|[<span data-ttu-id="c400f-629">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-629">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-630">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-630">ReadItem</span></span>|
|[<span data-ttu-id="c400f-631">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-631">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-632">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-632">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c400f-633">Exemple</span><span class="sxs-lookup"><span data-stu-id="c400f-633">Example</span></span>

```javascript
var seriesId = Office.context.mailbox.item.seriesId;
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="c400f-634">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="c400f-634">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="c400f-635">Obtient ou définit la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="c400f-635">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="c400f-p130">La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) pour convertir la valeur à la date et à l’heure du client.</span><span class="sxs-lookup"><span data-stu-id="c400f-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c400f-638">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-638">Read mode</span></span>

<span data-ttu-id="c400f-639">La propriété `start` renvoie un objet `Date`.</span><span class="sxs-lookup"><span data-stu-id="c400f-639">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c400f-640">Mode composition</span><span class="sxs-lookup"><span data-stu-id="c400f-640">Compose mode</span></span>

<span data-ttu-id="c400f-641">La propriété `start` renvoie un objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="c400f-641">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="c400f-642">Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.</span><span class="sxs-lookup"><span data-stu-id="c400f-642">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="c400f-643">Type :</span><span class="sxs-lookup"><span data-stu-id="c400f-643">Type:</span></span>

*   <span data-ttu-id="c400f-644">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="c400f-644">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c400f-645">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-645">Requirements</span></span>

|<span data-ttu-id="c400f-646">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-646">Requirement</span></span>|<span data-ttu-id="c400f-647">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-647">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-648">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-648">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-649">1.0</span><span class="sxs-lookup"><span data-stu-id="c400f-649">1.0</span></span>|
|[<span data-ttu-id="c400f-650">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-650">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-651">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-651">ReadItem</span></span>|
|[<span data-ttu-id="c400f-652">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-652">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-653">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-653">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c400f-654">Exemple</span><span class="sxs-lookup"><span data-stu-id="c400f-654">Example</span></span>

<span data-ttu-id="c400f-655">L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) de l’objet `Time`.</span><span class="sxs-lookup"><span data-stu-id="c400f-655">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```javascript
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

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="c400f-656">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="c400f-656">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="c400f-657">Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="c400f-657">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="c400f-658">La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.</span><span class="sxs-lookup"><span data-stu-id="c400f-658">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c400f-659">Mode lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-659">Read mode</span></span>

<span data-ttu-id="c400f-p131">La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.</span><span class="sxs-lookup"><span data-stu-id="c400f-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="c400f-662">Mode composition</span><span class="sxs-lookup"><span data-stu-id="c400f-662">Compose mode</span></span>

<span data-ttu-id="c400f-663">La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.</span><span class="sxs-lookup"><span data-stu-id="c400f-663">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c400f-664">Type :</span><span class="sxs-lookup"><span data-stu-id="c400f-664">Type:</span></span>

*   <span data-ttu-id="c400f-665">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="c400f-665">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c400f-666">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-666">Requirements</span></span>

|<span data-ttu-id="c400f-667">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-667">Requirement</span></span>|<span data-ttu-id="c400f-668">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-668">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-669">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-669">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-670">1.0</span><span class="sxs-lookup"><span data-stu-id="c400f-670">1.0</span></span>|
|[<span data-ttu-id="c400f-671">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-671">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-672">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-672">ReadItem</span></span>|
|[<span data-ttu-id="c400f-673">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-673">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-674">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-674">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="c400f-675">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c400f-675">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="c400f-676">Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message.</span><span class="sxs-lookup"><span data-stu-id="c400f-676">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="c400f-677">Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="c400f-677">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c400f-678">mode Lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-678">Read mode</span></span>

<span data-ttu-id="c400f-p133">La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. La collection est limitée à 100 membres.</span><span class="sxs-lookup"><span data-stu-id="c400f-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c400f-681">Mode composition</span><span class="sxs-lookup"><span data-stu-id="c400f-681">Compose mode</span></span>

<span data-ttu-id="c400f-682">La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message.</span><span class="sxs-lookup"><span data-stu-id="c400f-682">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="c400f-683">Type :</span><span class="sxs-lookup"><span data-stu-id="c400f-683">Type:</span></span>

*   <span data-ttu-id="c400f-684">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c400f-684">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c400f-685">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-685">Requirements</span></span>

|<span data-ttu-id="c400f-686">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-686">Requirement</span></span>|<span data-ttu-id="c400f-687">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-687">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-688">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-688">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-689">1.0</span><span class="sxs-lookup"><span data-stu-id="c400f-689">1.0</span></span>|
|[<span data-ttu-id="c400f-690">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-690">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-691">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-691">ReadItem</span></span>|
|[<span data-ttu-id="c400f-692">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-692">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-693">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-693">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c400f-694">Exemple</span><span class="sxs-lookup"><span data-stu-id="c400f-694">Example</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="c400f-695">Méthodes</span><span class="sxs-lookup"><span data-stu-id="c400f-695">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="c400f-696">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c400f-696">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="c400f-697">Ajoute un fichier à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="c400f-697">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="c400f-698">La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="c400f-698">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="c400f-699">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="c400f-699">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c400f-700">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="c400f-700">Parameters:</span></span>
|<span data-ttu-id="c400f-701">Nom</span><span class="sxs-lookup"><span data-stu-id="c400f-701">Name</span></span>|<span data-ttu-id="c400f-702">Type</span><span class="sxs-lookup"><span data-stu-id="c400f-702">Type</span></span>|<span data-ttu-id="c400f-703">Attributs</span><span class="sxs-lookup"><span data-stu-id="c400f-703">Attributes</span></span>|<span data-ttu-id="c400f-704">Description</span><span class="sxs-lookup"><span data-stu-id="c400f-704">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="c400f-705">String</span><span class="sxs-lookup"><span data-stu-id="c400f-705">String</span></span>||<span data-ttu-id="c400f-p134">URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.</span><span class="sxs-lookup"><span data-stu-id="c400f-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="c400f-708">String</span><span class="sxs-lookup"><span data-stu-id="c400f-708">String</span></span>||<span data-ttu-id="c400f-p135">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="c400f-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="c400f-711">Objet</span><span class="sxs-lookup"><span data-stu-id="c400f-711">Object</span></span>|<span data-ttu-id="c400f-712">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-712">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-713">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="c400f-713">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c400f-714">Objet</span><span class="sxs-lookup"><span data-stu-id="c400f-714">Object</span></span>|<span data-ttu-id="c400f-715">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-715">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-716">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="c400f-716">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="c400f-717">Boolean</span><span class="sxs-lookup"><span data-stu-id="c400f-717">Boolean</span></span>|<span data-ttu-id="c400f-718">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-718">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-719">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="c400f-719">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="c400f-720">fonction</span><span class="sxs-lookup"><span data-stu-id="c400f-720">function</span></span>|<span data-ttu-id="c400f-721">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-721">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-722">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c400f-722">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c400f-723">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="c400f-723">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="c400f-724">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="c400f-724">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c400f-725">Erreurs</span><span class="sxs-lookup"><span data-stu-id="c400f-725">Errors</span></span>

|<span data-ttu-id="c400f-726">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="c400f-726">Error code</span></span>|<span data-ttu-id="c400f-727">Description</span><span class="sxs-lookup"><span data-stu-id="c400f-727">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="c400f-728">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="c400f-728">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="c400f-729">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="c400f-729">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="c400f-730">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="c400f-730">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c400f-731">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-731">Requirements</span></span>

|<span data-ttu-id="c400f-732">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-732">Requirement</span></span>|<span data-ttu-id="c400f-733">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-733">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-734">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-734">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-735">1.1</span><span class="sxs-lookup"><span data-stu-id="c400f-735">1.1</span></span>|
|[<span data-ttu-id="c400f-736">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-736">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-737">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c400f-737">ReadWriteItem</span></span>|
|[<span data-ttu-id="c400f-738">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-738">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-739">Composition</span><span class="sxs-lookup"><span data-stu-id="c400f-739">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="c400f-740">Exemples</span><span class="sxs-lookup"><span data-stu-id="c400f-740">Examples</span></span>

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

<span data-ttu-id="c400f-741">L’exemple suivant montre comment ajouter un fichier image comme pièce jointe incorporée et comment la pièce jointe est affichée dans le corps du message.</span><span class="sxs-lookup"><span data-stu-id="c400f-741">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="c400f-742">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c400f-742">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="c400f-743">Ajoute un fichier provenant du codage base64 à un message ou un rendez-vous en pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="c400f-743">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="c400f-744">La méthode `addFileAttachmentFromBase64Async` charge le fichier depuis le codage base64 et le joint à l’élément dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="c400f-744">The `addFileAttachmentFromBase64Async` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span> <span data-ttu-id="c400f-745">Cette méthode renvoie l’identificateur de pièce jointe dans l’objet AsyncResult.value.</span><span class="sxs-lookup"><span data-stu-id="c400f-745">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="c400f-746">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="c400f-746">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c400f-747">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="c400f-747">Parameters:</span></span>
|<span data-ttu-id="c400f-748">Nom</span><span class="sxs-lookup"><span data-stu-id="c400f-748">Name</span></span>|<span data-ttu-id="c400f-749">Type</span><span class="sxs-lookup"><span data-stu-id="c400f-749">Type</span></span>|<span data-ttu-id="c400f-750">Attributs</span><span class="sxs-lookup"><span data-stu-id="c400f-750">Attributes</span></span>|<span data-ttu-id="c400f-751">Description</span><span class="sxs-lookup"><span data-stu-id="c400f-751">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="c400f-752">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c400f-752">String</span></span>||<span data-ttu-id="c400f-753">Contenu codé en base64 d’une image ou d’un fichier à ajouter à un e-mail ou à un événement.</span><span class="sxs-lookup"><span data-stu-id="c400f-753">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="c400f-754">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c400f-754">String</span></span>||<span data-ttu-id="c400f-p137">Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="c400f-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="c400f-757">Objet</span><span class="sxs-lookup"><span data-stu-id="c400f-757">Object</span></span>|<span data-ttu-id="c400f-758">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-758">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-759">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="c400f-759">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c400f-760">Objet</span><span class="sxs-lookup"><span data-stu-id="c400f-760">Object</span></span>|<span data-ttu-id="c400f-761">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-761">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-762">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="c400f-762">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="c400f-763">Boolean</span><span class="sxs-lookup"><span data-stu-id="c400f-763">Boolean</span></span>|<span data-ttu-id="c400f-764">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-764">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-765">Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="c400f-765">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="c400f-766">fonction</span><span class="sxs-lookup"><span data-stu-id="c400f-766">function</span></span>|<span data-ttu-id="c400f-767">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-767">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-768">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c400f-768">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c400f-769">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="c400f-769">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="c400f-770">En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="c400f-770">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c400f-771">Erreurs</span><span class="sxs-lookup"><span data-stu-id="c400f-771">Errors</span></span>

|<span data-ttu-id="c400f-772">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="c400f-772">Error code</span></span>|<span data-ttu-id="c400f-773">Description</span><span class="sxs-lookup"><span data-stu-id="c400f-773">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="c400f-774">La pièce jointe dépasse la taille autorisée.</span><span class="sxs-lookup"><span data-stu-id="c400f-774">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="c400f-775">La pièce jointe comporte une extension qui n’est pas autorisée.</span><span class="sxs-lookup"><span data-stu-id="c400f-775">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="c400f-776">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="c400f-776">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c400f-777">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-777">Requirements</span></span>

|<span data-ttu-id="c400f-778">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-778">Requirement</span></span>|<span data-ttu-id="c400f-779">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-779">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-780">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-780">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-781">Aperçu</span><span class="sxs-lookup"><span data-stu-id="c400f-781">Preview</span></span>|
|[<span data-ttu-id="c400f-782">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-782">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-783">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c400f-783">ReadWriteItem</span></span>|
|[<span data-ttu-id="c400f-784">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-784">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-785">Composition</span><span class="sxs-lookup"><span data-stu-id="c400f-785">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="c400f-786">Exemples</span><span class="sxs-lookup"><span data-stu-id="c400f-786">Examples</span></span>

```js
Office.context.mailbox.item.addFileAttachmentFromBase64Async(
  base64String,
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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="c400f-787">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c400f-787">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="c400f-788">Ajoute un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="c400f-788">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="c400f-789">Pour l’instant, les types d’événement pris en charge sont `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged` et `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="c400f-789">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c400f-790">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="c400f-790">Parameters:</span></span>

| <span data-ttu-id="c400f-791">Nom</span><span class="sxs-lookup"><span data-stu-id="c400f-791">Name</span></span> | <span data-ttu-id="c400f-792">Type</span><span class="sxs-lookup"><span data-stu-id="c400f-792">Type</span></span> | <span data-ttu-id="c400f-793">Attributs</span><span class="sxs-lookup"><span data-stu-id="c400f-793">Attributes</span></span> | <span data-ttu-id="c400f-794">Description</span><span class="sxs-lookup"><span data-stu-id="c400f-794">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="c400f-795">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="c400f-795">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="c400f-796">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="c400f-796">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="c400f-797">Fonction</span><span class="sxs-lookup"><span data-stu-id="c400f-797">Function</span></span> || <span data-ttu-id="c400f-p138">Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="c400f-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="c400f-801">Objet</span><span class="sxs-lookup"><span data-stu-id="c400f-801">Object</span></span> | <span data-ttu-id="c400f-802">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-802">&lt;optional&gt;</span></span> | <span data-ttu-id="c400f-803">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="c400f-803">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="c400f-804">Objet</span><span class="sxs-lookup"><span data-stu-id="c400f-804">Object</span></span> | <span data-ttu-id="c400f-805">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-805">&lt;optional&gt;</span></span> | <span data-ttu-id="c400f-806">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="c400f-806">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="c400f-807">fonction</span><span class="sxs-lookup"><span data-stu-id="c400f-807">function</span></span>| <span data-ttu-id="c400f-808">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-808">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-809">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c400f-809">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c400f-810">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-810">Requirements</span></span>

|<span data-ttu-id="c400f-811">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-811">Requirement</span></span>| <span data-ttu-id="c400f-812">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-812">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-813">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-813">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c400f-814">1.7</span><span class="sxs-lookup"><span data-stu-id="c400f-814">-17</span></span> |
|[<span data-ttu-id="c400f-815">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-815">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c400f-816">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-816">ReadItem</span></span> |
|[<span data-ttu-id="c400f-817">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-817">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c400f-818">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-818">Compose or read</span></span> |

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="c400f-819">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c400f-819">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="c400f-820">Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="c400f-820">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="c400f-p139">La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="c400f-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="c400f-824">L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="c400f-824">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="c400f-825">Si votre complément Office est exécuté dans Outlook Web App, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.</span><span class="sxs-lookup"><span data-stu-id="c400f-825">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c400f-826">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="c400f-826">Parameters:</span></span>

|<span data-ttu-id="c400f-827">Nom</span><span class="sxs-lookup"><span data-stu-id="c400f-827">Name</span></span>|<span data-ttu-id="c400f-828">Type</span><span class="sxs-lookup"><span data-stu-id="c400f-828">Type</span></span>|<span data-ttu-id="c400f-829">Attributs</span><span class="sxs-lookup"><span data-stu-id="c400f-829">Attributes</span></span>|<span data-ttu-id="c400f-830">Description</span><span class="sxs-lookup"><span data-stu-id="c400f-830">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="c400f-831">String</span><span class="sxs-lookup"><span data-stu-id="c400f-831">String</span></span>||<span data-ttu-id="c400f-p140">Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="c400f-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="c400f-834">String</span><span class="sxs-lookup"><span data-stu-id="c400f-834">String</span></span>||<span data-ttu-id="c400f-p141">Objet de l’élément à joindre. La longueur maximale est de 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="c400f-p141">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="c400f-837">Objet</span><span class="sxs-lookup"><span data-stu-id="c400f-837">Object</span></span>|<span data-ttu-id="c400f-838">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-838">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-839">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="c400f-839">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c400f-840">Objet</span><span class="sxs-lookup"><span data-stu-id="c400f-840">Object</span></span>|<span data-ttu-id="c400f-841">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-841">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-842">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="c400f-842">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c400f-843">fonction</span><span class="sxs-lookup"><span data-stu-id="c400f-843">function</span></span>|<span data-ttu-id="c400f-844">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-844">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-845">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c400f-845">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c400f-846">En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="c400f-846">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="c400f-847">En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.</span><span class="sxs-lookup"><span data-stu-id="c400f-847">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c400f-848">Erreurs</span><span class="sxs-lookup"><span data-stu-id="c400f-848">Errors</span></span>

|<span data-ttu-id="c400f-849">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="c400f-849">Error code</span></span>|<span data-ttu-id="c400f-850">Description</span><span class="sxs-lookup"><span data-stu-id="c400f-850">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="c400f-851">Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="c400f-851">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c400f-852">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-852">Requirements</span></span>

|<span data-ttu-id="c400f-853">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-853">Requirement</span></span>|<span data-ttu-id="c400f-854">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-854">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-855">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-855">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-856">1.1</span><span class="sxs-lookup"><span data-stu-id="c400f-856">1.1</span></span>|
|[<span data-ttu-id="c400f-857">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-857">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-858">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c400f-858">ReadWriteItem</span></span>|
|[<span data-ttu-id="c400f-859">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-859">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-860">Composition</span><span class="sxs-lookup"><span data-stu-id="c400f-860">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c400f-861">Exemple</span><span class="sxs-lookup"><span data-stu-id="c400f-861">Example</span></span>

<span data-ttu-id="c400f-862">L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.</span><span class="sxs-lookup"><span data-stu-id="c400f-862">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```javascript
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

####  <a name="close"></a><span data-ttu-id="c400f-863">close()</span><span class="sxs-lookup"><span data-stu-id="c400f-863">close()</span></span>

<span data-ttu-id="c400f-864">Ferme l’élément en cours qui est composé.</span><span class="sxs-lookup"><span data-stu-id="c400f-864">Closes the current item that is being composed.</span></span>

<span data-ttu-id="c400f-p142">Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.</span><span class="sxs-lookup"><span data-stu-id="c400f-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="c400f-867">Dans Outlook sur le web, si l’élément est un rendez-vous et s’il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, abandonner ou annuler même si aucune modification n’a été apportée depuis le dernier enregistrement de l’élément.</span><span class="sxs-lookup"><span data-stu-id="c400f-867">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="c400f-868">Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.</span><span class="sxs-lookup"><span data-stu-id="c400f-868">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c400f-869">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-869">Requirements</span></span>

|<span data-ttu-id="c400f-870">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-870">Requirement</span></span>|<span data-ttu-id="c400f-871">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-871">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-872">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-872">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-873">1.3</span><span class="sxs-lookup"><span data-stu-id="c400f-873">1.3</span></span>|
|[<span data-ttu-id="c400f-874">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-874">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-875">Restreinte</span><span class="sxs-lookup"><span data-stu-id="c400f-875">Restricted</span></span>|
|[<span data-ttu-id="c400f-876">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-876">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-877">Composition</span><span class="sxs-lookup"><span data-stu-id="c400f-877">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="c400f-878">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="c400f-878">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="c400f-879">Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="c400f-879">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c400f-880">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="c400f-880">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c400f-881">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="c400f-881">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="c400f-882">Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="c400f-882">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="c400f-p143">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="c400f-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c400f-886">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="c400f-886">Parameters:</span></span>

|<span data-ttu-id="c400f-887">Nom</span><span class="sxs-lookup"><span data-stu-id="c400f-887">Name</span></span>|<span data-ttu-id="c400f-888">Type</span><span class="sxs-lookup"><span data-stu-id="c400f-888">Type</span></span>|<span data-ttu-id="c400f-889">Attributs</span><span class="sxs-lookup"><span data-stu-id="c400f-889">Attributes</span></span>|<span data-ttu-id="c400f-890">Description</span><span class="sxs-lookup"><span data-stu-id="c400f-890">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="c400f-891">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="c400f-891">String &#124; Object</span></span>||<span data-ttu-id="c400f-p144">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="c400f-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="c400f-894">**OU**</span><span class="sxs-lookup"><span data-stu-id="c400f-894">**OR**</span></span><br/><span data-ttu-id="c400f-p145">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="c400f-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="c400f-897">String</span><span class="sxs-lookup"><span data-stu-id="c400f-897">String</span></span>|<span data-ttu-id="c400f-898">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-898">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-p146">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="c400f-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="c400f-901">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-901">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="c400f-902">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-902">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-903">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="c400f-903">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="c400f-904">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c400f-904">String</span></span>||<span data-ttu-id="c400f-p147">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="c400f-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="c400f-907">String</span><span class="sxs-lookup"><span data-stu-id="c400f-907">String</span></span>||<span data-ttu-id="c400f-908">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="c400f-908">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="c400f-909">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c400f-909">String</span></span>||<span data-ttu-id="c400f-p148">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="c400f-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="c400f-912">Booléen</span><span class="sxs-lookup"><span data-stu-id="c400f-912">Boolean</span></span>||<span data-ttu-id="c400f-p149">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="c400f-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="c400f-915">String</span><span class="sxs-lookup"><span data-stu-id="c400f-915">String</span></span>||<span data-ttu-id="c400f-p150">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="c400f-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="c400f-919">function</span><span class="sxs-lookup"><span data-stu-id="c400f-919">function</span></span>|<span data-ttu-id="c400f-920">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-920">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-921">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c400f-921">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c400f-922">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-922">Requirements</span></span>

|<span data-ttu-id="c400f-923">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-923">Requirement</span></span>|<span data-ttu-id="c400f-924">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-924">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-925">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-925">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-926">1.0</span><span class="sxs-lookup"><span data-stu-id="c400f-926">1.0</span></span>|
|[<span data-ttu-id="c400f-927">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-927">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-928">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-928">ReadItem</span></span>|
|[<span data-ttu-id="c400f-929">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-929">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-930">Lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-930">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="c400f-931">Exemples</span><span class="sxs-lookup"><span data-stu-id="c400f-931">Examples</span></span>

<span data-ttu-id="c400f-932">Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.</span><span class="sxs-lookup"><span data-stu-id="c400f-932">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="c400f-933">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="c400f-933">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="c400f-934">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="c400f-934">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="c400f-935">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="c400f-935">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="c400f-936">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="c400f-936">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="c400f-937">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="c400f-937">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="c400f-938">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="c400f-938">displayReplyForm(formData)</span></span>

<span data-ttu-id="c400f-939">Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.</span><span class="sxs-lookup"><span data-stu-id="c400f-939">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c400f-940">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="c400f-940">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c400f-941">Dans Outlook Web App, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.</span><span class="sxs-lookup"><span data-stu-id="c400f-941">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="c400f-942">Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.</span><span class="sxs-lookup"><span data-stu-id="c400f-942">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="c400f-p151">Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook Web App tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.</span><span class="sxs-lookup"><span data-stu-id="c400f-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c400f-946">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="c400f-946">Parameters:</span></span>

|<span data-ttu-id="c400f-947">Nom</span><span class="sxs-lookup"><span data-stu-id="c400f-947">Name</span></span>|<span data-ttu-id="c400f-948">Type</span><span class="sxs-lookup"><span data-stu-id="c400f-948">Type</span></span>|<span data-ttu-id="c400f-949">Attributs</span><span class="sxs-lookup"><span data-stu-id="c400f-949">Attributes</span></span>|<span data-ttu-id="c400f-950">Description</span><span class="sxs-lookup"><span data-stu-id="c400f-950">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="c400f-951">String &#124; Object</span><span class="sxs-lookup"><span data-stu-id="c400f-951">String &#124; Object</span></span>||<span data-ttu-id="c400f-p152">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="c400f-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="c400f-954">**OU**</span><span class="sxs-lookup"><span data-stu-id="c400f-954">**OR**</span></span><br/><span data-ttu-id="c400f-p153">Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :</span><span class="sxs-lookup"><span data-stu-id="c400f-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="c400f-957">String</span><span class="sxs-lookup"><span data-stu-id="c400f-957">String</span></span>|<span data-ttu-id="c400f-958">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-958">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-p154">Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="c400f-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="c400f-961">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-961">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="c400f-962">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-962">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-963">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="c400f-963">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="c400f-964">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c400f-964">String</span></span>||<span data-ttu-id="c400f-p155">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="c400f-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="c400f-967">String</span><span class="sxs-lookup"><span data-stu-id="c400f-967">String</span></span>||<span data-ttu-id="c400f-968">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="c400f-968">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="c400f-969">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c400f-969">String</span></span>||<span data-ttu-id="c400f-p156">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="c400f-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="c400f-972">Booléen</span><span class="sxs-lookup"><span data-stu-id="c400f-972">Boolean</span></span>||<span data-ttu-id="c400f-p157">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="c400f-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="c400f-975">String</span><span class="sxs-lookup"><span data-stu-id="c400f-975">String</span></span>||<span data-ttu-id="c400f-p158">Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="c400f-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="c400f-979">function</span><span class="sxs-lookup"><span data-stu-id="c400f-979">function</span></span>|<span data-ttu-id="c400f-980">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-980">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-981">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c400f-981">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c400f-982">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-982">Requirements</span></span>

|<span data-ttu-id="c400f-983">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-983">Requirement</span></span>|<span data-ttu-id="c400f-984">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-984">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-985">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-985">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-986">1.0</span><span class="sxs-lookup"><span data-stu-id="c400f-986">1.0</span></span>|
|[<span data-ttu-id="c400f-987">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-987">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-988">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-988">ReadItem</span></span>|
|[<span data-ttu-id="c400f-989">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-989">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-990">Lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-990">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="c400f-991">Exemples</span><span class="sxs-lookup"><span data-stu-id="c400f-991">Examples</span></span>

<span data-ttu-id="c400f-992">Le code suivant transmet une chaîne à la fonction `displayReplyForm`.</span><span class="sxs-lookup"><span data-stu-id="c400f-992">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="c400f-993">Réponse avec un corps vide.</span><span class="sxs-lookup"><span data-stu-id="c400f-993">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="c400f-994">Réponse avec un corps.</span><span class="sxs-lookup"><span data-stu-id="c400f-994">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="c400f-995">Réponse avec un corps et la pièce jointe d’un fichier.</span><span class="sxs-lookup"><span data-stu-id="c400f-995">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="c400f-996">Réponse avec un corps et la pièce jointe d’un élément.</span><span class="sxs-lookup"><span data-stu-id="c400f-996">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="c400f-997">Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.</span><span class="sxs-lookup"><span data-stu-id="c400f-997">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="c400f-998">getAttachmentContentAsync(attachmentId, [options], callback) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="c400f-998">getAttachmentContentAsync(attachmentId, [options], callback) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="c400f-999">Permet d’obtenir la pièce jointe spécifiée depuis un message ou un rendez-vous, et la renvoie en tant qu’objet `AttachmentContent`.</span><span class="sxs-lookup"><span data-stu-id="c400f-999">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="c400f-1000">La méthode `getAttachmentContentAsync` permet d’obtenir la pièce jointe avec l’identificateur spécifié de l’élément.</span><span class="sxs-lookup"><span data-stu-id="c400f-1000">The `getAttachmentContentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="c400f-1001">Nous vous recommandons de suivre la bonne pratique consistant à utiliser l’identificateur pour récupérer une pièce jointe dans la même session que celle où les objets attachmentIds ont été récupérés avec l’appel `getAttachmentsAsync` ou `item.attachments`.</span><span class="sxs-lookup"><span data-stu-id="c400f-1001">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="c400f-1002">Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session.</span><span class="sxs-lookup"><span data-stu-id="c400f-1002">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="c400f-1003">Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer un formulaire incorporé qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="c400f-1003">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c400f-1004">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="c400f-1004">Parameters:</span></span>

|<span data-ttu-id="c400f-1005">Nom</span><span class="sxs-lookup"><span data-stu-id="c400f-1005">Name</span></span>|<span data-ttu-id="c400f-1006">Type</span><span class="sxs-lookup"><span data-stu-id="c400f-1006">Type</span></span>|<span data-ttu-id="c400f-1007">Attributs</span><span class="sxs-lookup"><span data-stu-id="c400f-1007">Attributes</span></span>|<span data-ttu-id="c400f-1008">Description</span><span class="sxs-lookup"><span data-stu-id="c400f-1008">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="c400f-1009">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c400f-1009">String</span></span>||<span data-ttu-id="c400f-1010">Identificateur de la pièce jointe que vous voulez obtenir.</span><span class="sxs-lookup"><span data-stu-id="c400f-1010">The identifier of the attachment you want to get.</span></span> <span data-ttu-id="c400f-1011">La longueur maximale de la chaîne est 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="c400f-1011">The maximum length of the string is 100 characters.</span></span>|
|`options`|<span data-ttu-id="c400f-1012">Object</span><span class="sxs-lookup"><span data-stu-id="c400f-1012">Object</span></span>|<span data-ttu-id="c400f-1013">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-1013">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-1014">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="c400f-1014">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c400f-1015">Objet</span><span class="sxs-lookup"><span data-stu-id="c400f-1015">Object</span></span>|<span data-ttu-id="c400f-1016">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-1016">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-1017">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="c400f-1017">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c400f-1018">fonction</span><span class="sxs-lookup"><span data-stu-id="c400f-1018">function</span></span>|<span data-ttu-id="c400f-1019">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-1019">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-1020">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c400f-1020">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c400f-1021">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-1021">Requirements</span></span>

|<span data-ttu-id="c400f-1022">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-1022">Requirement</span></span>|<span data-ttu-id="c400f-1023">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-1023">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-1024">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-1024">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-1025">Aperçu</span><span class="sxs-lookup"><span data-stu-id="c400f-1025">Preview</span></span>|
|[<span data-ttu-id="c400f-1026">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-1026">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-1027">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-1027">ReadItem</span></span>|
|[<span data-ttu-id="c400f-1028">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-1028">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-1029">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-1029">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c400f-1030">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="c400f-1030">Returns:</span></span>

<span data-ttu-id="c400f-1031">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="c400f-1031">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="c400f-1032">Exemple</span><span class="sxs-lookup"><span data-stu-id="c400f-1032">Example</span></span>

```javascript
var item = Office.context.mailbox.item;
var listOfAttachments = [];
item.getAttachmentsAsync(callback);
function callback(result) {
    if (result.value.length > 0) {
        for (i = 0 ; i < result.value.length ; i++) {
            var options = {asyncContext: {type: result.value[i].attachmentType}};
            getAttachmentContentAsync(result.value[i].id, options, handleAttachmentsCallback);  
        }
    }
}

function handleAttachmentsCallback(result) {
    // parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file
    if (result.format == Office.MailboxEnums.AttachmentContentFormat.Base64) {
        // handle file attachment
    }
    else if (result.format == Office.MailboxEnums.AttachmentContentFormat.Eml) {
        // handle item attachment
    }
    else if (result.format == Office.MailboxEnums.AttachmentContentFormat.ICalendar) {
        // handle .icalender attachment
    }
    else {
        // handle cloud attachment  
    }
}
```

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="c400f-1033">getAttachmentsAsync([options], callback) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="c400f-1033">getAttachmentsAsync([options], callback) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="c400f-1034">Permet d’obtenir les pièces jointes de l’élément sous forme de tableau.</span><span class="sxs-lookup"><span data-stu-id="c400f-1034">Gets the item's attachments as an array.</span></span> <span data-ttu-id="c400f-1035">Mode composition uniquement.</span><span class="sxs-lookup"><span data-stu-id="c400f-1035">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c400f-1036">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="c400f-1036">Parameters:</span></span>

|<span data-ttu-id="c400f-1037">Nom</span><span class="sxs-lookup"><span data-stu-id="c400f-1037">Name</span></span>|<span data-ttu-id="c400f-1038">Type</span><span class="sxs-lookup"><span data-stu-id="c400f-1038">Type</span></span>|<span data-ttu-id="c400f-1039">Attributs</span><span class="sxs-lookup"><span data-stu-id="c400f-1039">Attributes</span></span>|<span data-ttu-id="c400f-1040">Description</span><span class="sxs-lookup"><span data-stu-id="c400f-1040">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="c400f-1041">Objet</span><span class="sxs-lookup"><span data-stu-id="c400f-1041">Object</span></span>|<span data-ttu-id="c400f-1042">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-1042">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-1043">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="c400f-1043">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c400f-1044">Objet</span><span class="sxs-lookup"><span data-stu-id="c400f-1044">Object</span></span>|<span data-ttu-id="c400f-1045">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-1045">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-1046">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="c400f-1046">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c400f-1047">fonction</span><span class="sxs-lookup"><span data-stu-id="c400f-1047">function</span></span>|<span data-ttu-id="c400f-1048">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-1048">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-1049">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c400f-1049">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c400f-1050">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-1050">Requirements</span></span>

|<span data-ttu-id="c400f-1051">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-1051">Requirement</span></span>|<span data-ttu-id="c400f-1052">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-1052">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-1053">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-1053">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-1054">Aperçu</span><span class="sxs-lookup"><span data-stu-id="c400f-1054">Preview</span></span>|
|[<span data-ttu-id="c400f-1055">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-1055">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-1056">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-1056">ReadItem</span></span>|
|[<span data-ttu-id="c400f-1057">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-1057">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-1058">Composition</span><span class="sxs-lookup"><span data-stu-id="c400f-1058">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="c400f-1059">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="c400f-1059">Returns:</span></span>

<span data-ttu-id="c400f-1060">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="c400f-1060">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="c400f-1061">Exemple</span><span class="sxs-lookup"><span data-stu-id="c400f-1061">Example</span></span>

<span data-ttu-id="c400f-1062">L’exemple suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="c400f-1062">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```javascript
var item = Office.context.mailbox.item;
var outputString = "";
item.getAttachmentsAsync(callback);  
function callback(result) {
    if (result.value.length > 0) {
        for (i = 0 ; i < result.value.length ; i++) {
            var _att = result.value [i];
            outputString += "<BR>" + i + ". Name: ";
            outputString += _att.name;
            outputString += "<BR>ID: " + _att.id;
            outputString += "<BR>contentType: " + _att.contentType;
            outputString += "<BR>size: " + _att.size;
            outputString += "<BR>attachmentType: " + _att.attachmentType;
            outputString += "<BR>isInline: " + _att.isInline;
        }
    }
}
```

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="c400f-1063">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="c400f-1063">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="c400f-1064">Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="c400f-1064">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="c400f-1065">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="c400f-1065">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c400f-1066">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-1066">Requirements</span></span>

|<span data-ttu-id="c400f-1067">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-1067">Requirement</span></span>|<span data-ttu-id="c400f-1068">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-1068">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-1069">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-1069">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-1070">1.0</span><span class="sxs-lookup"><span data-stu-id="c400f-1070">1.0</span></span>|
|[<span data-ttu-id="c400f-1071">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-1071">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-1072">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-1072">ReadItem</span></span>|
|[<span data-ttu-id="c400f-1073">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-1073">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-1074">Lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-1074">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c400f-1075">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="c400f-1075">Returns:</span></span>

<span data-ttu-id="c400f-1076">Type : [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="c400f-1076">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="c400f-1077">Exemple</span><span class="sxs-lookup"><span data-stu-id="c400f-1077">Example</span></span>

<span data-ttu-id="c400f-1078">L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="c400f-1078">The following example accesses the contacts entities on the current item.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="c400f-1079">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="c400f-1079">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="c400f-1080">Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="c400f-1080">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="c400f-1081">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="c400f-1081">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c400f-1082">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="c400f-1082">Parameters:</span></span>

|<span data-ttu-id="c400f-1083">Nom</span><span class="sxs-lookup"><span data-stu-id="c400f-1083">Name</span></span>|<span data-ttu-id="c400f-1084">Type</span><span class="sxs-lookup"><span data-stu-id="c400f-1084">Type</span></span>|<span data-ttu-id="c400f-1085">Description</span><span class="sxs-lookup"><span data-stu-id="c400f-1085">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="c400f-1086">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="c400f-1086">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="c400f-1087">Une des valeurs d’énumération EntityType.</span><span class="sxs-lookup"><span data-stu-id="c400f-1087">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c400f-1088">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-1088">Requirements</span></span>

|<span data-ttu-id="c400f-1089">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-1089">Requirement</span></span>|<span data-ttu-id="c400f-1090">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-1090">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-1091">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-1091">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-1092">1.0</span><span class="sxs-lookup"><span data-stu-id="c400f-1092">1.0</span></span>|
|[<span data-ttu-id="c400f-1093">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-1093">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-1094">Restreinte</span><span class="sxs-lookup"><span data-stu-id="c400f-1094">Restricted</span></span>|
|[<span data-ttu-id="c400f-1095">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-1095">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-1096">Lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-1096">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c400f-1097">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="c400f-1097">Returns:</span></span>

<span data-ttu-id="c400f-1098">Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null.</span><span class="sxs-lookup"><span data-stu-id="c400f-1098">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="c400f-1099">Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="c400f-1099">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="c400f-1100">Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.</span><span class="sxs-lookup"><span data-stu-id="c400f-1100">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="c400f-1101">Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.</span><span class="sxs-lookup"><span data-stu-id="c400f-1101">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="c400f-1102">Valeur de `entityType`</span><span class="sxs-lookup"><span data-stu-id="c400f-1102">Value of `entityType`</span></span>|<span data-ttu-id="c400f-1103">Type des objets du tableau renvoyé</span><span class="sxs-lookup"><span data-stu-id="c400f-1103">Type of objects in returned array</span></span>|<span data-ttu-id="c400f-1104">Niveau d’autorisation requis</span><span class="sxs-lookup"><span data-stu-id="c400f-1104">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="c400f-1105">String</span><span class="sxs-lookup"><span data-stu-id="c400f-1105">String</span></span>|<span data-ttu-id="c400f-1106">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="c400f-1106">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="c400f-1107">Contact</span><span class="sxs-lookup"><span data-stu-id="c400f-1107">Contact</span></span>|<span data-ttu-id="c400f-1108">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c400f-1108">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="c400f-1109">String</span><span class="sxs-lookup"><span data-stu-id="c400f-1109">String</span></span>|<span data-ttu-id="c400f-1110">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c400f-1110">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="c400f-1111">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="c400f-1111">MeetingSuggestion</span></span>|<span data-ttu-id="c400f-1112">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c400f-1112">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="c400f-1113">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="c400f-1113">PhoneNumber</span></span>|<span data-ttu-id="c400f-1114">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="c400f-1114">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="c400f-1115">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="c400f-1115">TaskSuggestion</span></span>|<span data-ttu-id="c400f-1116">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c400f-1116">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="c400f-1117">String</span><span class="sxs-lookup"><span data-stu-id="c400f-1117">String</span></span>|<span data-ttu-id="c400f-1118">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="c400f-1118">**Restricted**</span></span>|

<span data-ttu-id="c400f-1119">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="c400f-1119">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="c400f-1120">Exemple</span><span class="sxs-lookup"><span data-stu-id="c400f-1120">Example</span></span>

<span data-ttu-id="c400f-1121">L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="c400f-1121">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="c400f-1122">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="c400f-1122">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="c400f-1123">Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="c400f-1123">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c400f-1124">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="c400f-1124">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c400f-1125">La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="c400f-1125">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c400f-1126">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="c400f-1126">Parameters:</span></span>

|<span data-ttu-id="c400f-1127">Nom</span><span class="sxs-lookup"><span data-stu-id="c400f-1127">Name</span></span>|<span data-ttu-id="c400f-1128">Type</span><span class="sxs-lookup"><span data-stu-id="c400f-1128">Type</span></span>|<span data-ttu-id="c400f-1129">Description</span><span class="sxs-lookup"><span data-stu-id="c400f-1129">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="c400f-1130">String</span><span class="sxs-lookup"><span data-stu-id="c400f-1130">String</span></span>|<span data-ttu-id="c400f-1131">Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="c400f-1131">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c400f-1132">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-1132">Requirements</span></span>

|<span data-ttu-id="c400f-1133">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-1133">Requirement</span></span>|<span data-ttu-id="c400f-1134">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-1134">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-1135">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-1135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-1136">1.0</span><span class="sxs-lookup"><span data-stu-id="c400f-1136">1.0</span></span>|
|[<span data-ttu-id="c400f-1137">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-1137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-1138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-1138">ReadItem</span></span>|
|[<span data-ttu-id="c400f-1139">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-1139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-1140">Lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-1140">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c400f-1141">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="c400f-1141">Returns:</span></span>

<span data-ttu-id="c400f-p163">Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.</span><span class="sxs-lookup"><span data-stu-id="c400f-p163">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="c400f-1144">Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="c400f-1144">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="c400f-1145">getInitializationContextAsync([options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c400f-1145">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="c400f-1146">Récupère les données d’initialisation transmises quand le complément est [activé par un message actionnable](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="c400f-1146">Gets initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="c400f-1147">Cette méthode est uniquement prise en charge par Outlook 2016 ou version ultérieure pour Windows (versions en un clic supérieures à 16.0.8413.1000) et Outlook sur le web pour Office 365.</span><span class="sxs-lookup"><span data-stu-id="c400f-1147">Note: This method is only supported by Outlook 2016 for Windows (Click-to-Run versions greater than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c400f-1148">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="c400f-1148">Parameters:</span></span>
|<span data-ttu-id="c400f-1149">Nom</span><span class="sxs-lookup"><span data-stu-id="c400f-1149">Name</span></span>|<span data-ttu-id="c400f-1150">Type</span><span class="sxs-lookup"><span data-stu-id="c400f-1150">Type</span></span>|<span data-ttu-id="c400f-1151">Attributs</span><span class="sxs-lookup"><span data-stu-id="c400f-1151">Attributes</span></span>|<span data-ttu-id="c400f-1152">Description</span><span class="sxs-lookup"><span data-stu-id="c400f-1152">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="c400f-1153">Objet</span><span class="sxs-lookup"><span data-stu-id="c400f-1153">Object</span></span>|<span data-ttu-id="c400f-1154">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-1154">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-1155">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="c400f-1155">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c400f-1156">Objet</span><span class="sxs-lookup"><span data-stu-id="c400f-1156">Object</span></span>|<span data-ttu-id="c400f-1157">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-1157">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-1158">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="c400f-1158">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c400f-1159">fonction</span><span class="sxs-lookup"><span data-stu-id="c400f-1159">function</span></span>|<span data-ttu-id="c400f-1160">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-1160">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-1161">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c400f-1161">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c400f-1162">En cas de réussite, les données d’initialisation sont fournies dans la propriété `asyncResult.value` sous forme de chaîne.</span><span class="sxs-lookup"><span data-stu-id="c400f-1162">On success, the intialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="c400f-1163">S’il n’existe aucun contexte d’initialisation, l’objet `asyncResult` contient un objet `Error` dont la propriété `code` est définie sur `9020` et la propriété `name` sur `GenericResponseError`.</span><span class="sxs-lookup"><span data-stu-id="c400f-1163">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c400f-1164">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-1164">Requirements</span></span>

|<span data-ttu-id="c400f-1165">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-1165">Requirement</span></span>|<span data-ttu-id="c400f-1166">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-1166">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-1167">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-1167">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-1168">Aperçu</span><span class="sxs-lookup"><span data-stu-id="c400f-1168">Preview</span></span>|
|[<span data-ttu-id="c400f-1169">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-1169">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-1170">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-1170">ReadItem</span></span>|
|[<span data-ttu-id="c400f-1171">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-1171">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-1172">Lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-1172">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c400f-1173">Exemple</span><span class="sxs-lookup"><span data-stu-id="c400f-1173">Example</span></span>

```javascript
// Get the initialization context (if present)
Office.context.mailbox.item.getInitializationContextAsync(
  function(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      if (asyncResult.value != null && asyncResult.value.length > 0) {
        // The value is a string, parse to an object
        var context = JSON.parse(asyncResult.value);
        // Do something with context
      } else {
        // Empty context, treat as no context
      }
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is
        // no context
        // Treat as no context
      } else {
        // Handle the error
      }
    }
  }
);
```

#### <a name="getregexmatches--object"></a><span data-ttu-id="c400f-1174">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="c400f-1174">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="c400f-1175">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="c400f-1175">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c400f-1176">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="c400f-1176">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c400f-p164">La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="c400f-p164">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="c400f-1180">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="c400f-1180">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="c400f-1181">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="c400f-1181">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="c400f-p165">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="c400f-p165">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c400f-1185">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-1185">Requirements</span></span>

|<span data-ttu-id="c400f-1186">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-1186">Requirement</span></span>|<span data-ttu-id="c400f-1187">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-1187">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-1188">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-1188">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-1189">1.0</span><span class="sxs-lookup"><span data-stu-id="c400f-1189">1.0</span></span>|
|[<span data-ttu-id="c400f-1190">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-1190">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-1191">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-1191">ReadItem</span></span>|
|[<span data-ttu-id="c400f-1192">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-1192">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-1193">Lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-1193">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c400f-1194">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="c400f-1194">Returns:</span></span>

<span data-ttu-id="c400f-p166">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="c400f-p166">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="c400f-1197">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="c400f-1197">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c400f-1198">Object</span><span class="sxs-lookup"><span data-stu-id="c400f-1198">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="c400f-1199">Exemple</span><span class="sxs-lookup"><span data-stu-id="c400f-1199">Example</span></span>

<span data-ttu-id="c400f-1200">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="c400f-1200">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="c400f-1201">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="c400f-1201">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="c400f-1202">Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="c400f-1202">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c400f-1203">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="c400f-1203">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c400f-1204">La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.</span><span class="sxs-lookup"><span data-stu-id="c400f-1204">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="c400f-p167">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.</span><span class="sxs-lookup"><span data-stu-id="c400f-p167">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c400f-1207">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="c400f-1207">Parameters:</span></span>

|<span data-ttu-id="c400f-1208">Nom</span><span class="sxs-lookup"><span data-stu-id="c400f-1208">Name</span></span>|<span data-ttu-id="c400f-1209">Type</span><span class="sxs-lookup"><span data-stu-id="c400f-1209">Type</span></span>|<span data-ttu-id="c400f-1210">Description</span><span class="sxs-lookup"><span data-stu-id="c400f-1210">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="c400f-1211">String</span><span class="sxs-lookup"><span data-stu-id="c400f-1211">String</span></span>|<span data-ttu-id="c400f-1212">Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.</span><span class="sxs-lookup"><span data-stu-id="c400f-1212">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c400f-1213">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-1213">Requirements</span></span>

|<span data-ttu-id="c400f-1214">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-1214">Requirement</span></span>|<span data-ttu-id="c400f-1215">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-1215">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-1216">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-1216">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-1217">1.0</span><span class="sxs-lookup"><span data-stu-id="c400f-1217">1.0</span></span>|
|[<span data-ttu-id="c400f-1218">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-1218">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-1219">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-1219">ReadItem</span></span>|
|[<span data-ttu-id="c400f-1220">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-1220">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-1221">Lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-1221">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c400f-1222">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="c400f-1222">Returns:</span></span>

<span data-ttu-id="c400f-1223">Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.</span><span class="sxs-lookup"><span data-stu-id="c400f-1223">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="c400f-1224">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="c400f-1224">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c400f-1225">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="c400f-1225">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="c400f-1226">Exemple</span><span class="sxs-lookup"><span data-stu-id="c400f-1226">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="c400f-1227">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="c400f-1227">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="c400f-1228">Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="c400f-1228">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="c400f-p168">Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.</span><span class="sxs-lookup"><span data-stu-id="c400f-p168">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c400f-1231">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="c400f-1231">Parameters:</span></span>

|<span data-ttu-id="c400f-1232">Nom</span><span class="sxs-lookup"><span data-stu-id="c400f-1232">Name</span></span>|<span data-ttu-id="c400f-1233">Type</span><span class="sxs-lookup"><span data-stu-id="c400f-1233">Type</span></span>|<span data-ttu-id="c400f-1234">Attributs</span><span class="sxs-lookup"><span data-stu-id="c400f-1234">Attributes</span></span>|<span data-ttu-id="c400f-1235">Description</span><span class="sxs-lookup"><span data-stu-id="c400f-1235">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="c400f-1236">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="c400f-1236">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="c400f-p169">Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.</span><span class="sxs-lookup"><span data-stu-id="c400f-p169">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="c400f-1240">Object</span><span class="sxs-lookup"><span data-stu-id="c400f-1240">Object</span></span>|<span data-ttu-id="c400f-1241">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-1241">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-1242">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="c400f-1242">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c400f-1243">Objet</span><span class="sxs-lookup"><span data-stu-id="c400f-1243">Object</span></span>|<span data-ttu-id="c400f-1244">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-1244">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-1245">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="c400f-1245">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c400f-1246">fonction</span><span class="sxs-lookup"><span data-stu-id="c400f-1246">function</span></span>||<span data-ttu-id="c400f-1247">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c400f-1247">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c400f-1248">Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`.</span><span class="sxs-lookup"><span data-stu-id="c400f-1248">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="c400f-1249">Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.</span><span class="sxs-lookup"><span data-stu-id="c400f-1249">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c400f-1250">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-1250">Requirements</span></span>

|<span data-ttu-id="c400f-1251">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-1251">Requirement</span></span>|<span data-ttu-id="c400f-1252">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-1252">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-1253">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-1253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-1254">1.2</span><span class="sxs-lookup"><span data-stu-id="c400f-1254">1.2</span></span>|
|[<span data-ttu-id="c400f-1255">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-1255">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-1256">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c400f-1256">ReadWriteItem</span></span>|
|[<span data-ttu-id="c400f-1257">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-1257">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-1258">Composition</span><span class="sxs-lookup"><span data-stu-id="c400f-1258">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="c400f-1259">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="c400f-1259">Returns:</span></span>

<span data-ttu-id="c400f-1260">Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.</span><span class="sxs-lookup"><span data-stu-id="c400f-1260">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="c400f-1261">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="c400f-1261">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c400f-1262">String</span><span class="sxs-lookup"><span data-stu-id="c400f-1262">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="c400f-1263">Exemple</span><span class="sxs-lookup"><span data-stu-id="c400f-1263">Example</span></span>

```javascript
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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="c400f-1264">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="c400f-1264">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="c400f-p171">Permet d’obtenir les entités figurant dans une correspondance en surbrillance qu’un utilisateur a sélectionné. Les correspondances en surbrillance s’appliquent aux [compléments contextuels](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="c400f-p171">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="c400f-1267">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="c400f-1267">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c400f-1268">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-1268">Requirements</span></span>

|<span data-ttu-id="c400f-1269">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-1269">Requirement</span></span>|<span data-ttu-id="c400f-1270">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-1270">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-1271">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-1271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-1272">1.6</span><span class="sxs-lookup"><span data-stu-id="c400f-1272">-16</span></span>|
|[<span data-ttu-id="c400f-1273">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-1273">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-1274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-1274">ReadItem</span></span>|
|[<span data-ttu-id="c400f-1275">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-1275">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-1276">Lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-1276">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c400f-1277">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="c400f-1277">Returns:</span></span>

<span data-ttu-id="c400f-1278">Type : [Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="c400f-1278">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="c400f-1279">Exemple</span><span class="sxs-lookup"><span data-stu-id="c400f-1279">Example</span></span>

<span data-ttu-id="c400f-1280">L’exemple suivant accède aux entités d’adresses dans la correspondance en surbrillance sélectionnée par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="c400f-1280">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="c400f-1281">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="c400f-1281">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="c400f-p172">Renvoie des valeurs de chaîne dans une correspondance en surbrillance, qui correspondent aux expressions régulières définies dans le fichier manifeste XML. Les correspondances en surbrillance s’appliquent aux [compléments contextuels](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span><span class="sxs-lookup"><span data-stu-id="c400f-p172">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="c400f-1284">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="c400f-1284">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c400f-p173">La méthode `getSelectedRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.</span><span class="sxs-lookup"><span data-stu-id="c400f-p173">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="c400f-1288">Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :</span><span class="sxs-lookup"><span data-stu-id="c400f-1288">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="c400f-1289">L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.</span><span class="sxs-lookup"><span data-stu-id="c400f-1289">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="c400f-p174">Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.</span><span class="sxs-lookup"><span data-stu-id="c400f-p174">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c400f-1293">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-1293">Requirements</span></span>

|<span data-ttu-id="c400f-1294">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-1294">Requirement</span></span>|<span data-ttu-id="c400f-1295">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-1295">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-1296">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-1296">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-1297">1.6</span><span class="sxs-lookup"><span data-stu-id="c400f-1297">-16</span></span>|
|[<span data-ttu-id="c400f-1298">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-1298">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-1299">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-1299">ReadItem</span></span>|
|[<span data-ttu-id="c400f-1300">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-1300">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-1301">Lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-1301">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c400f-1302">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="c400f-1302">Returns:</span></span>

<span data-ttu-id="c400f-p175">Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.</span><span class="sxs-lookup"><span data-stu-id="c400f-p175">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="c400f-1305">Exemple</span><span class="sxs-lookup"><span data-stu-id="c400f-1305">Example</span></span>

<span data-ttu-id="c400f-1306">L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.</span><span class="sxs-lookup"><span data-stu-id="c400f-1306">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="c400f-1307">getSharedPropertiesAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="c400f-1307">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="c400f-1308">Permet d’obtenir les propriétés du rendez-vous ou du message sélectionné dans une boîte aux lettres, un calendrier ou un dossier partagé.</span><span class="sxs-lookup"><span data-stu-id="c400f-1308">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c400f-1309">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="c400f-1309">Parameters:</span></span>

|<span data-ttu-id="c400f-1310">Nom</span><span class="sxs-lookup"><span data-stu-id="c400f-1310">Name</span></span>|<span data-ttu-id="c400f-1311">Type</span><span class="sxs-lookup"><span data-stu-id="c400f-1311">Type</span></span>|<span data-ttu-id="c400f-1312">Attributs</span><span class="sxs-lookup"><span data-stu-id="c400f-1312">Attributes</span></span>|<span data-ttu-id="c400f-1313">Description</span><span class="sxs-lookup"><span data-stu-id="c400f-1313">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="c400f-1314">Objet</span><span class="sxs-lookup"><span data-stu-id="c400f-1314">Object</span></span>|<span data-ttu-id="c400f-1315">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-1315">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-1316">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="c400f-1316">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c400f-1317">Objet</span><span class="sxs-lookup"><span data-stu-id="c400f-1317">Object</span></span>|<span data-ttu-id="c400f-1318">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-1318">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-1319">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="c400f-1319">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c400f-1320">fonction</span><span class="sxs-lookup"><span data-stu-id="c400f-1320">function</span></span>||<span data-ttu-id="c400f-1321">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c400f-1321">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c400f-1322">Les propriétés partagées sont fournies sous la forme d’un objet [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="c400f-1322">The custom properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="c400f-1323">Cet objet peut être utilisé pour obtenir des propriétés partagées de l’élément.</span><span class="sxs-lookup"><span data-stu-id="c400f-1323">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c400f-1324">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-1324">Requirements</span></span>

|<span data-ttu-id="c400f-1325">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-1325">Requirement</span></span>|<span data-ttu-id="c400f-1326">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-1326">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-1327">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-1327">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-1328">Aperçu</span><span class="sxs-lookup"><span data-stu-id="c400f-1328">Preview</span></span>|
|[<span data-ttu-id="c400f-1329">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-1329">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-1330">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-1330">ReadItem</span></span>|
|[<span data-ttu-id="c400f-1331">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-1331">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-1332">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-1332">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c400f-1333">Exemple</span><span class="sxs-lookup"><span data-stu-id="c400f-1333">Example</span></span>

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);
function callback (asyncResult) {
  var context=asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="c400f-1334">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="c400f-1334">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="c400f-1335">Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="c400f-1335">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="c400f-p177">Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.</span><span class="sxs-lookup"><span data-stu-id="c400f-p177">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c400f-1339">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="c400f-1339">Parameters:</span></span>

|<span data-ttu-id="c400f-1340">Nom</span><span class="sxs-lookup"><span data-stu-id="c400f-1340">Name</span></span>|<span data-ttu-id="c400f-1341">Type</span><span class="sxs-lookup"><span data-stu-id="c400f-1341">Type</span></span>|<span data-ttu-id="c400f-1342">Attributs</span><span class="sxs-lookup"><span data-stu-id="c400f-1342">Attributes</span></span>|<span data-ttu-id="c400f-1343">Description</span><span class="sxs-lookup"><span data-stu-id="c400f-1343">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="c400f-1344">function</span><span class="sxs-lookup"><span data-stu-id="c400f-1344">function</span></span>||<span data-ttu-id="c400f-1345">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c400f-1345">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c400f-1346">Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook/office.customproperties) dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="c400f-1346">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="c400f-1347">Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.</span><span class="sxs-lookup"><span data-stu-id="c400f-1347">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="c400f-1348">Objet</span><span class="sxs-lookup"><span data-stu-id="c400f-1348">Object</span></span>|<span data-ttu-id="c400f-1349">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-1349">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-1350">Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="c400f-1350">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="c400f-1351">Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="c400f-1351">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c400f-1352">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-1352">Requirements</span></span>

|<span data-ttu-id="c400f-1353">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-1353">Requirement</span></span>|<span data-ttu-id="c400f-1354">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-1354">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-1355">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-1355">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-1356">1.0</span><span class="sxs-lookup"><span data-stu-id="c400f-1356">1.0</span></span>|
|[<span data-ttu-id="c400f-1357">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-1357">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-1358">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-1358">ReadItem</span></span>|
|[<span data-ttu-id="c400f-1359">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-1359">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-1360">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-1360">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c400f-1361">Exemple</span><span class="sxs-lookup"><span data-stu-id="c400f-1361">Example</span></span>

<span data-ttu-id="c400f-p180">L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.</span><span class="sxs-lookup"><span data-stu-id="c400f-p180">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```javascript
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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="c400f-1365">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c400f-1365">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="c400f-1366">Supprime une pièce jointe d’un message ou d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="c400f-1366">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="c400f-1367">La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément.</span><span class="sxs-lookup"><span data-stu-id="c400f-1367">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="c400f-1368">Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session.</span><span class="sxs-lookup"><span data-stu-id="c400f-1368">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="c400f-1369">Dans Outlook Web App et OWA pour les périphériques, l’identificateur de pièce jointe n’est valable que dans la même session.</span><span class="sxs-lookup"><span data-stu-id="c400f-1369">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="c400f-1370">Une session est terminée lorsque l’utilisateur ferme l’application, ou si celui-ci commence à composer un formulaire incorporé qu’il fait ensuite apparaître dans une fenêtre séparée.</span><span class="sxs-lookup"><span data-stu-id="c400f-1370">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c400f-1371">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="c400f-1371">Parameters:</span></span>

|<span data-ttu-id="c400f-1372">Nom</span><span class="sxs-lookup"><span data-stu-id="c400f-1372">Name</span></span>|<span data-ttu-id="c400f-1373">Type</span><span class="sxs-lookup"><span data-stu-id="c400f-1373">Type</span></span>|<span data-ttu-id="c400f-1374">Attributs</span><span class="sxs-lookup"><span data-stu-id="c400f-1374">Attributes</span></span>|<span data-ttu-id="c400f-1375">Description</span><span class="sxs-lookup"><span data-stu-id="c400f-1375">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="c400f-1376">String</span><span class="sxs-lookup"><span data-stu-id="c400f-1376">String</span></span>||<span data-ttu-id="c400f-p182">Identificateur de la pièce jointe à supprimer. La longueur maximale de la chaîne est de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="c400f-p182">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`|<span data-ttu-id="c400f-1379">Objet</span><span class="sxs-lookup"><span data-stu-id="c400f-1379">Object</span></span>|<span data-ttu-id="c400f-1380">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-1380">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-1381">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="c400f-1381">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c400f-1382">Objet</span><span class="sxs-lookup"><span data-stu-id="c400f-1382">Object</span></span>|<span data-ttu-id="c400f-1383">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-1383">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-1384">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="c400f-1384">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c400f-1385">fonction</span><span class="sxs-lookup"><span data-stu-id="c400f-1385">function</span></span>|<span data-ttu-id="c400f-1386">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-1386">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-1387">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c400f-1387">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c400f-1388">En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.</span><span class="sxs-lookup"><span data-stu-id="c400f-1388">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c400f-1389">Erreurs</span><span class="sxs-lookup"><span data-stu-id="c400f-1389">Errors</span></span>

|<span data-ttu-id="c400f-1390">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="c400f-1390">Error code</span></span>|<span data-ttu-id="c400f-1391">Description</span><span class="sxs-lookup"><span data-stu-id="c400f-1391">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="c400f-1392">L’identificateur de la pièce jointe n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="c400f-1392">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c400f-1393">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-1393">Requirements</span></span>

|<span data-ttu-id="c400f-1394">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-1394">Requirement</span></span>|<span data-ttu-id="c400f-1395">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-1395">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-1396">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-1396">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-1397">1.1</span><span class="sxs-lookup"><span data-stu-id="c400f-1397">1.1</span></span>|
|[<span data-ttu-id="c400f-1398">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-1398">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-1399">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c400f-1399">ReadWriteItem</span></span>|
|[<span data-ttu-id="c400f-1400">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-1400">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-1401">Composition</span><span class="sxs-lookup"><span data-stu-id="c400f-1401">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c400f-1402">Exemple</span><span class="sxs-lookup"><span data-stu-id="c400f-1402">Example</span></span>

<span data-ttu-id="c400f-1403">Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».</span><span class="sxs-lookup"><span data-stu-id="c400f-1403">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="c400f-1404">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c400f-1404">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="c400f-1405">Retire un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="c400f-1405">Removes an event handler for a</span></span>

<span data-ttu-id="c400f-1406">Pour l’instant, les types d’événement pris en charge sont `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged` et `Office.EventType.RecurrenceChanged`.</span><span class="sxs-lookup"><span data-stu-id="c400f-1406">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c400f-1407">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="c400f-1407">Parameters:</span></span>

| <span data-ttu-id="c400f-1408">Nom</span><span class="sxs-lookup"><span data-stu-id="c400f-1408">Name</span></span> | <span data-ttu-id="c400f-1409">Type</span><span class="sxs-lookup"><span data-stu-id="c400f-1409">Type</span></span> | <span data-ttu-id="c400f-1410">Attributs</span><span class="sxs-lookup"><span data-stu-id="c400f-1410">Attributes</span></span> | <span data-ttu-id="c400f-1411">Description</span><span class="sxs-lookup"><span data-stu-id="c400f-1411">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="c400f-1412">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="c400f-1412">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="c400f-1413">Événement qui doit révoquer le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="c400f-1413">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="c400f-1414">Fonction</span><span class="sxs-lookup"><span data-stu-id="c400f-1414">Function</span></span> || <span data-ttu-id="c400f-p183">Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `removeHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="c400f-p183">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `removeHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="c400f-1418">Objet</span><span class="sxs-lookup"><span data-stu-id="c400f-1418">Object</span></span> | <span data-ttu-id="c400f-1419">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-1419">&lt;optional&gt;</span></span> | <span data-ttu-id="c400f-1420">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="c400f-1420">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="c400f-1421">Objet</span><span class="sxs-lookup"><span data-stu-id="c400f-1421">Object</span></span> | <span data-ttu-id="c400f-1422">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-1422">&lt;optional&gt;</span></span> | <span data-ttu-id="c400f-1423">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="c400f-1423">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="c400f-1424">fonction</span><span class="sxs-lookup"><span data-stu-id="c400f-1424">function</span></span>| <span data-ttu-id="c400f-1425">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-1425">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-1426">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c400f-1426">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c400f-1427">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-1427">Requirements</span></span>

|<span data-ttu-id="c400f-1428">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-1428">Requirement</span></span>| <span data-ttu-id="c400f-1429">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-1429">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-1430">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-1430">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c400f-1431">1.7</span><span class="sxs-lookup"><span data-stu-id="c400f-1431">-17</span></span> |
|[<span data-ttu-id="c400f-1432">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-1432">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c400f-1433">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c400f-1433">ReadItem</span></span> |
|[<span data-ttu-id="c400f-1434">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-1434">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c400f-1435">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c400f-1435">Compose or read</span></span> |

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="c400f-1436">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="c400f-1436">saveAsync([options], callback)</span></span>

<span data-ttu-id="c400f-1437">Enregistre un élément de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="c400f-1437">Asynchronously saves an item.</span></span>

<span data-ttu-id="c400f-p184">Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel. Dans Outlook Web App ou Outlook en mode en ligne, l’élément est enregistré sur le serveur. Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.</span><span class="sxs-lookup"><span data-stu-id="c400f-p184">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="c400f-1441">si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` à utiliser avec EWS ou l’API REST, sachez que lorsqu’Outlook est en mode mis en cache, la synchronisation de l’élément sur le serveur peut prendre un certain temps.</span><span class="sxs-lookup"><span data-stu-id="c400f-1441">Note: If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the  will return an error.</span></span> <span data-ttu-id="c400f-1442">Avant que l’élément ne soit synchronisé, l’utilisation de l’élément `itemId` renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="c400f-1442">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="c400f-p186">Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.</span><span class="sxs-lookup"><span data-stu-id="c400f-p186">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="c400f-1446">Les clients suivants ont un comportement différent avec `saveAsync` sur les rendez-vous en mode composition :</span><span class="sxs-lookup"><span data-stu-id="c400f-1446">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="c400f-1447">Outlook pour Mac ne prend pas en charge `saveAsync` sur une réunion en mode composition.</span><span class="sxs-lookup"><span data-stu-id="c400f-1447">Note: Mac Outlook does not support `saveAsync` on a meeting in compose mode. Calling  on a meeting in Mac Outlook will return an error.</span></span> <span data-ttu-id="c400f-1448">Le fait d’appeler `saveAsync` sur une réunion dans Outlook pour Mac renvoie une erreur.</span><span class="sxs-lookup"><span data-stu-id="c400f-1448">Note: Mac Outlook does not support  on a meeting in compose mode. Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="c400f-1449">Outlook sur le web envoie toujours une invitation ou une mise à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.</span><span class="sxs-lookup"><span data-stu-id="c400f-1449">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c400f-1450">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="c400f-1450">Parameters:</span></span>

|<span data-ttu-id="c400f-1451">Nom</span><span class="sxs-lookup"><span data-stu-id="c400f-1451">Name</span></span>|<span data-ttu-id="c400f-1452">Type</span><span class="sxs-lookup"><span data-stu-id="c400f-1452">Type</span></span>|<span data-ttu-id="c400f-1453">Attributs</span><span class="sxs-lookup"><span data-stu-id="c400f-1453">Attributes</span></span>|<span data-ttu-id="c400f-1454">Description</span><span class="sxs-lookup"><span data-stu-id="c400f-1454">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="c400f-1455">Objet</span><span class="sxs-lookup"><span data-stu-id="c400f-1455">Object</span></span>|<span data-ttu-id="c400f-1456">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-1456">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-1457">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="c400f-1457">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c400f-1458">Objet</span><span class="sxs-lookup"><span data-stu-id="c400f-1458">Object</span></span>|<span data-ttu-id="c400f-1459">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-1459">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-1460">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="c400f-1460">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c400f-1461">fonction</span><span class="sxs-lookup"><span data-stu-id="c400f-1461">function</span></span>||<span data-ttu-id="c400f-1462">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c400f-1462">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c400f-1463">En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="c400f-1463">On success, the item identifier is provided in the `asyncResult.value` property.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c400f-1464">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-1464">Requirements</span></span>

|<span data-ttu-id="c400f-1465">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-1465">Requirement</span></span>|<span data-ttu-id="c400f-1466">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-1466">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-1467">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-1467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-1468">1.3</span><span class="sxs-lookup"><span data-stu-id="c400f-1468">1.3</span></span>|
|[<span data-ttu-id="c400f-1469">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-1469">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-1470">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c400f-1470">ReadWriteItem</span></span>|
|[<span data-ttu-id="c400f-1471">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-1471">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-1472">Composition</span><span class="sxs-lookup"><span data-stu-id="c400f-1472">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="c400f-1473">範例</span><span class="sxs-lookup"><span data-stu-id="c400f-1473">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="c400f-p188">Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.</span><span class="sxs-lookup"><span data-stu-id="c400f-p188">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="c400f-1476">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="c400f-1476">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="c400f-1477">Insère les données dans le corps ou l’objet d’un message de manière asynchrone.</span><span class="sxs-lookup"><span data-stu-id="c400f-1477">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="c400f-p189">La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.</span><span class="sxs-lookup"><span data-stu-id="c400f-p189">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c400f-1481">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="c400f-1481">Parameters:</span></span>

|<span data-ttu-id="c400f-1482">Nom</span><span class="sxs-lookup"><span data-stu-id="c400f-1482">Name</span></span>|<span data-ttu-id="c400f-1483">Type</span><span class="sxs-lookup"><span data-stu-id="c400f-1483">Type</span></span>|<span data-ttu-id="c400f-1484">Attributs</span><span class="sxs-lookup"><span data-stu-id="c400f-1484">Attributes</span></span>|<span data-ttu-id="c400f-1485">Description</span><span class="sxs-lookup"><span data-stu-id="c400f-1485">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="c400f-1486">String</span><span class="sxs-lookup"><span data-stu-id="c400f-1486">String</span></span>||<span data-ttu-id="c400f-p190">Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.</span><span class="sxs-lookup"><span data-stu-id="c400f-p190">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="c400f-1490">Objet</span><span class="sxs-lookup"><span data-stu-id="c400f-1490">Object</span></span>|<span data-ttu-id="c400f-1491">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-1491">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-1492">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="c400f-1492">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c400f-1493">Objet</span><span class="sxs-lookup"><span data-stu-id="c400f-1493">Object</span></span>|<span data-ttu-id="c400f-1494">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-1494">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-1495">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="c400f-1495">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="c400f-1496">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="c400f-1496">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="c400f-1497">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c400f-1497">&lt;optional&gt;</span></span>|<span data-ttu-id="c400f-p191">Si `text`, le style existant est appliqué dans Outlook Web App et Outlook. Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.</span><span class="sxs-lookup"><span data-stu-id="c400f-p191">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="c400f-p192">Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook Web App et le style par défaut dans Outlook. Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.</span><span class="sxs-lookup"><span data-stu-id="c400f-p192">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="c400f-1502">Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.</span><span class="sxs-lookup"><span data-stu-id="c400f-1502">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="c400f-1503">fonction</span><span class="sxs-lookup"><span data-stu-id="c400f-1503">function</span></span>||<span data-ttu-id="c400f-1504">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="c400f-1504">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c400f-1505">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c400f-1505">Requirements</span></span>

|<span data-ttu-id="c400f-1506">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c400f-1506">Requirement</span></span>|<span data-ttu-id="c400f-1507">Valeur</span><span class="sxs-lookup"><span data-stu-id="c400f-1507">Value</span></span>|
|---|---|
|[<span data-ttu-id="c400f-1508">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c400f-1508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c400f-1509">1.2</span><span class="sxs-lookup"><span data-stu-id="c400f-1509">1.2</span></span>|
|[<span data-ttu-id="c400f-1510">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c400f-1510">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c400f-1511">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c400f-1511">ReadWriteItem</span></span>|
|[<span data-ttu-id="c400f-1512">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c400f-1512">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c400f-1513">Composition</span><span class="sxs-lookup"><span data-stu-id="c400f-1513">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c400f-1514">Exemple</span><span class="sxs-lookup"><span data-stu-id="c400f-1514">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```