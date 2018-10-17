 

# <a name="office"></a><span data-ttu-id="8edc6-101">Office</span><span class="sxs-lookup"><span data-stu-id="8edc6-101">Office</span></span>

<span data-ttu-id="8edc6-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[Interface API partagée](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="8edc6-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="8edc6-104">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8edc6-104">Requirements</span></span>

|<span data-ttu-id="8edc6-105">Condition requise</span><span class="sxs-lookup"><span data-stu-id="8edc6-105">Requirement</span></span>| <span data-ttu-id="8edc6-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="8edc6-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="8edc6-107">Version minimale requise de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8edc6-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8edc6-108">1.0</span><span class="sxs-lookup"><span data-stu-id="8edc6-108">1.0</span></span>|
|[<span data-ttu-id="8edc6-109">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8edc6-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8edc6-110">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="8edc6-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="8edc6-111">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="8edc6-111">Members and methods</span></span>

| <span data-ttu-id="8edc6-112">Membre</span><span class="sxs-lookup"><span data-stu-id="8edc6-112">Member</span></span> | <span data-ttu-id="8edc6-113">Type</span><span class="sxs-lookup"><span data-stu-id="8edc6-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="8edc6-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="8edc6-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="8edc6-115">Membre</span><span class="sxs-lookup"><span data-stu-id="8edc6-115">Member</span></span> |
| [<span data-ttu-id="8edc6-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="8edc6-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="8edc6-117">Membre</span><span class="sxs-lookup"><span data-stu-id="8edc6-117">Member</span></span> |
| [<span data-ttu-id="8edc6-118">EventType</span><span class="sxs-lookup"><span data-stu-id="8edc6-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="8edc6-119">Membre</span><span class="sxs-lookup"><span data-stu-id="8edc6-119">Member</span></span> |
| [<span data-ttu-id="8edc6-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="8edc6-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="8edc6-121">Membre</span><span class="sxs-lookup"><span data-stu-id="8edc6-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="8edc6-122">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="8edc6-122">Namespaces</span></span>

<span data-ttu-id="8edc6-123">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="8edc6-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="8edc6-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="8edc6-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="8edc6-125">Membres</span><span class="sxs-lookup"><span data-stu-id="8edc6-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="8edc6-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="8edc6-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="8edc6-127">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="8edc6-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="8edc6-128">Type :</span><span class="sxs-lookup"><span data-stu-id="8edc6-128">Type:</span></span>

*   <span data-ttu-id="8edc6-129">String</span><span class="sxs-lookup"><span data-stu-id="8edc6-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8edc6-130">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="8edc6-130">Properties:</span></span>

|<span data-ttu-id="8edc6-131">Nom</span><span class="sxs-lookup"><span data-stu-id="8edc6-131">Name</span></span>| <span data-ttu-id="8edc6-132">Type</span><span class="sxs-lookup"><span data-stu-id="8edc6-132">Type</span></span>| <span data-ttu-id="8edc6-133">Description</span><span class="sxs-lookup"><span data-stu-id="8edc6-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="8edc6-134">String</span><span class="sxs-lookup"><span data-stu-id="8edc6-134">String</span></span>|<span data-ttu-id="8edc6-135">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="8edc6-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="8edc6-136">String</span><span class="sxs-lookup"><span data-stu-id="8edc6-136">String</span></span>|<span data-ttu-id="8edc6-137">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="8edc6-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8edc6-138">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8edc6-138">Requirements</span></span>

|<span data-ttu-id="8edc6-139">Condition requise</span><span class="sxs-lookup"><span data-stu-id="8edc6-139">Requirement</span></span>| <span data-ttu-id="8edc6-140">Valeur</span><span class="sxs-lookup"><span data-stu-id="8edc6-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="8edc6-141">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8edc6-141">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8edc6-142">1.0</span><span class="sxs-lookup"><span data-stu-id="8edc6-142">1.0</span></span>|
|[<span data-ttu-id="8edc6-143">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8edc6-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8edc6-144">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="8edc6-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="8edc6-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="8edc6-145">CoercionType :String</span></span>

<span data-ttu-id="8edc6-146">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="8edc6-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="8edc6-147">Type :</span><span class="sxs-lookup"><span data-stu-id="8edc6-147">Type:</span></span>

*   <span data-ttu-id="8edc6-148">String</span><span class="sxs-lookup"><span data-stu-id="8edc6-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8edc6-149">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="8edc6-149">Properties:</span></span>

|<span data-ttu-id="8edc6-150">Nom</span><span class="sxs-lookup"><span data-stu-id="8edc6-150">Name</span></span>| <span data-ttu-id="8edc6-151">Type</span><span class="sxs-lookup"><span data-stu-id="8edc6-151">Type</span></span>| <span data-ttu-id="8edc6-152">Description</span><span class="sxs-lookup"><span data-stu-id="8edc6-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="8edc6-153">String</span><span class="sxs-lookup"><span data-stu-id="8edc6-153">String</span></span>|<span data-ttu-id="8edc6-154">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="8edc6-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="8edc6-155">String</span><span class="sxs-lookup"><span data-stu-id="8edc6-155">String</span></span>|<span data-ttu-id="8edc6-156">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="8edc6-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8edc6-157">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8edc6-157">Requirements</span></span>

|<span data-ttu-id="8edc6-158">Condition requise</span><span class="sxs-lookup"><span data-stu-id="8edc6-158">Requirement</span></span>| <span data-ttu-id="8edc6-159">Valeur</span><span class="sxs-lookup"><span data-stu-id="8edc6-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="8edc6-160">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8edc6-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8edc6-161">1.0</span><span class="sxs-lookup"><span data-stu-id="8edc6-161">1.0</span></span>|
|[<span data-ttu-id="8edc6-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8edc6-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8edc6-163">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="8edc6-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="8edc6-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="8edc6-164">EventType :String</span></span>

<span data-ttu-id="8edc6-165">Spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="8edc6-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="8edc6-166">Type :</span><span class="sxs-lookup"><span data-stu-id="8edc6-166">Type:</span></span>

*   <span data-ttu-id="8edc6-167">String</span><span class="sxs-lookup"><span data-stu-id="8edc6-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8edc6-168">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="8edc6-168">Properties:</span></span>

| <span data-ttu-id="8edc6-169">Nom</span><span class="sxs-lookup"><span data-stu-id="8edc6-169">Name</span></span> | <span data-ttu-id="8edc6-170">Type</span><span class="sxs-lookup"><span data-stu-id="8edc6-170">Type</span></span> | <span data-ttu-id="8edc6-171">Description</span><span class="sxs-lookup"><span data-stu-id="8edc6-171">Description</span></span> | <span data-ttu-id="8edc6-172">Ensemble minimal de conditions requises</span><span class="sxs-lookup"><span data-stu-id="8edc6-172">Minimum mailbox requirement set version</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="8edc6-173">String</span><span class="sxs-lookup"><span data-stu-id="8edc6-173">String</span></span> | <span data-ttu-id="8edc6-174">La date ou l’heure du rendez-vous sélectionné ou de la série a changé.</span><span class="sxs-lookup"><span data-stu-id="8edc6-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="8edc6-175">1.7</span><span class="sxs-lookup"><span data-stu-id="8edc6-175">-17</span></span> |
|`ItemChanged`| <span data-ttu-id="8edc6-176">String</span><span class="sxs-lookup"><span data-stu-id="8edc6-176">String</span></span> | <span data-ttu-id="8edc6-177">L’élément sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="8edc6-177">The selected item has changed.</span></span> | <span data-ttu-id="8edc6-178">1.5</span><span class="sxs-lookup"><span data-stu-id="8edc6-178">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="8edc6-179">String</span><span class="sxs-lookup"><span data-stu-id="8edc6-179">String</span></span> | <span data-ttu-id="8edc6-180">L’élément sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="8edc6-180">The selected item has changed.</span></span> | <span data-ttu-id="8edc6-181">Aperçu</span><span class="sxs-lookup"><span data-stu-id="8edc6-181">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="8edc6-182">String</span><span class="sxs-lookup"><span data-stu-id="8edc6-182">String</span></span> | <span data-ttu-id="8edc6-183">La liste de destinataires de l’élément sélectionné ou l’emplacement du rendez-vous a changé.</span><span class="sxs-lookup"><span data-stu-id="8edc6-183">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="8edc6-184">1.7</span><span class="sxs-lookup"><span data-stu-id="8edc6-184">-17</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="8edc6-185">String</span><span class="sxs-lookup"><span data-stu-id="8edc6-185">String</span></span> | <span data-ttu-id="8edc6-186">La périodicité de la série sélectionnée a changé.</span><span class="sxs-lookup"><span data-stu-id="8edc6-186">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="8edc6-187">1.7</span><span class="sxs-lookup"><span data-stu-id="8edc6-187">-17</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8edc6-188">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8edc6-188">Requirements</span></span>

|<span data-ttu-id="8edc6-189">Condition requise</span><span class="sxs-lookup"><span data-stu-id="8edc6-189">Requirement</span></span>| <span data-ttu-id="8edc6-190">Valeur</span><span class="sxs-lookup"><span data-stu-id="8edc6-190">Value</span></span>|
|---|---|
|[<span data-ttu-id="8edc6-191">Version minimale de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8edc6-191">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8edc6-192">1,5</span><span class="sxs-lookup"><span data-stu-id="8edc6-192">1.5</span></span> |
|[<span data-ttu-id="8edc6-193">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8edc6-193">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8edc6-194">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="8edc6-194">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="8edc6-195">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="8edc6-195">SourceProperty :String</span></span>

<span data-ttu-id="8edc6-196">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="8edc6-196">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="8edc6-197">Type :</span><span class="sxs-lookup"><span data-stu-id="8edc6-197">Type:</span></span>

*   <span data-ttu-id="8edc6-198">String</span><span class="sxs-lookup"><span data-stu-id="8edc6-198">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8edc6-199">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="8edc6-199">Properties:</span></span>

|<span data-ttu-id="8edc6-200">Nom</span><span class="sxs-lookup"><span data-stu-id="8edc6-200">Name</span></span>| <span data-ttu-id="8edc6-201">Type</span><span class="sxs-lookup"><span data-stu-id="8edc6-201">Type</span></span>| <span data-ttu-id="8edc6-202">Description</span><span class="sxs-lookup"><span data-stu-id="8edc6-202">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="8edc6-203">String</span><span class="sxs-lookup"><span data-stu-id="8edc6-203">String</span></span>|<span data-ttu-id="8edc6-204">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="8edc6-204">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="8edc6-205">String</span><span class="sxs-lookup"><span data-stu-id="8edc6-205">String</span></span>|<span data-ttu-id="8edc6-206">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="8edc6-206">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8edc6-207">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8edc6-207">Requirements</span></span>

|<span data-ttu-id="8edc6-208">Condition requise</span><span class="sxs-lookup"><span data-stu-id="8edc6-208">Requirement</span></span>| <span data-ttu-id="8edc6-209">Valeur</span><span class="sxs-lookup"><span data-stu-id="8edc6-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="8edc6-210">Version minimale requise de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8edc6-210">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8edc6-211">1.0</span><span class="sxs-lookup"><span data-stu-id="8edc6-211">1.0</span></span>|
|[<span data-ttu-id="8edc6-212">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8edc6-212">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8edc6-213">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="8edc6-213">Compose or read</span></span>|