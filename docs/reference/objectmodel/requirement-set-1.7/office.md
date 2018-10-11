 

# <a name="office"></a><span data-ttu-id="eda72-101">Office</span><span class="sxs-lookup"><span data-stu-id="eda72-101">Office</span></span>

<span data-ttu-id="eda72-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[Interface API partagée](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="eda72-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="eda72-104">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="eda72-104">Requirements</span></span>

|<span data-ttu-id="eda72-105">Condition</span><span class="sxs-lookup"><span data-stu-id="eda72-105">Requirement</span></span>| <span data-ttu-id="eda72-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="eda72-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="eda72-107">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="eda72-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eda72-108">1.0</span><span class="sxs-lookup"><span data-stu-id="eda72-108">1.0</span></span>|
|[<span data-ttu-id="eda72-109">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="eda72-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="eda72-110">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="eda72-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="eda72-111">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="eda72-111">Members and methods</span></span>

| <span data-ttu-id="eda72-112">Membre</span><span class="sxs-lookup"><span data-stu-id="eda72-112">Member</span></span> | <span data-ttu-id="eda72-113">Type</span><span class="sxs-lookup"><span data-stu-id="eda72-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="eda72-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="eda72-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="eda72-115">Membre</span><span class="sxs-lookup"><span data-stu-id="eda72-115">Member</span></span> |
| [<span data-ttu-id="eda72-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="eda72-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="eda72-117">Membre</span><span class="sxs-lookup"><span data-stu-id="eda72-117">Member</span></span> |
| [<span data-ttu-id="eda72-118">EventType</span><span class="sxs-lookup"><span data-stu-id="eda72-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="eda72-119">Membre</span><span class="sxs-lookup"><span data-stu-id="eda72-119">Member</span></span> |
| [<span data-ttu-id="eda72-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="eda72-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="eda72-121">Membre</span><span class="sxs-lookup"><span data-stu-id="eda72-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="eda72-122">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="eda72-122">Namespaces</span></span>

<span data-ttu-id="eda72-123">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="eda72-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="eda72-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="eda72-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="eda72-125">Membres</span><span class="sxs-lookup"><span data-stu-id="eda72-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="eda72-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="eda72-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="eda72-127">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="eda72-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="eda72-128">Type :</span><span class="sxs-lookup"><span data-stu-id="eda72-128">Type:</span></span>

*   <span data-ttu-id="eda72-129">String</span><span class="sxs-lookup"><span data-stu-id="eda72-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="eda72-130">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="eda72-130">Properties:</span></span>

|<span data-ttu-id="eda72-131">Nom</span><span class="sxs-lookup"><span data-stu-id="eda72-131">Name</span></span>| <span data-ttu-id="eda72-132">Type</span><span class="sxs-lookup"><span data-stu-id="eda72-132">Type</span></span>| <span data-ttu-id="eda72-133">Description</span><span class="sxs-lookup"><span data-stu-id="eda72-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="eda72-134">String</span><span class="sxs-lookup"><span data-stu-id="eda72-134">String</span></span>|<span data-ttu-id="eda72-135">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="eda72-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="eda72-136">String</span><span class="sxs-lookup"><span data-stu-id="eda72-136">String</span></span>|<span data-ttu-id="eda72-137">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="eda72-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="eda72-138">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="eda72-138">Requirements</span></span>

|<span data-ttu-id="eda72-139">Condition</span><span class="sxs-lookup"><span data-stu-id="eda72-139">Requirement</span></span>| <span data-ttu-id="eda72-140">Valeur</span><span class="sxs-lookup"><span data-stu-id="eda72-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="eda72-141">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="eda72-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eda72-142">1.0</span><span class="sxs-lookup"><span data-stu-id="eda72-142">1.0</span></span>|
|[<span data-ttu-id="eda72-143">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="eda72-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="eda72-144">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="eda72-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="eda72-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="eda72-145">CoercionType :String</span></span>

<span data-ttu-id="eda72-146">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="eda72-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="eda72-147">Type :</span><span class="sxs-lookup"><span data-stu-id="eda72-147">Type:</span></span>

*   <span data-ttu-id="eda72-148">String</span><span class="sxs-lookup"><span data-stu-id="eda72-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="eda72-149">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="eda72-149">Properties:</span></span>

|<span data-ttu-id="eda72-150">Nom</span><span class="sxs-lookup"><span data-stu-id="eda72-150">Name</span></span>| <span data-ttu-id="eda72-151">Type</span><span class="sxs-lookup"><span data-stu-id="eda72-151">Type</span></span>| <span data-ttu-id="eda72-152">Description</span><span class="sxs-lookup"><span data-stu-id="eda72-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="eda72-153">String</span><span class="sxs-lookup"><span data-stu-id="eda72-153">String</span></span>|<span data-ttu-id="eda72-154">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="eda72-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="eda72-155">String</span><span class="sxs-lookup"><span data-stu-id="eda72-155">String</span></span>|<span data-ttu-id="eda72-156">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="eda72-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="eda72-157">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="eda72-157">Requirements</span></span>

|<span data-ttu-id="eda72-158">Condition</span><span class="sxs-lookup"><span data-stu-id="eda72-158">Requirement</span></span>| <span data-ttu-id="eda72-159">Valeur</span><span class="sxs-lookup"><span data-stu-id="eda72-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="eda72-160">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="eda72-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eda72-161">1.0</span><span class="sxs-lookup"><span data-stu-id="eda72-161">1.0</span></span>|
|[<span data-ttu-id="eda72-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="eda72-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="eda72-163">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="eda72-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="eda72-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="eda72-164">EventType :String</span></span>

<span data-ttu-id="eda72-165">Spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="eda72-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="eda72-166">Type :</span><span class="sxs-lookup"><span data-stu-id="eda72-166">Type:</span></span>

*   <span data-ttu-id="eda72-167">String</span><span class="sxs-lookup"><span data-stu-id="eda72-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="eda72-168">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="eda72-168">Properties:</span></span>

| <span data-ttu-id="eda72-169">Nom</span><span class="sxs-lookup"><span data-stu-id="eda72-169">Name</span></span> | <span data-ttu-id="eda72-170">Type</span><span class="sxs-lookup"><span data-stu-id="eda72-170">Type</span></span> | <span data-ttu-id="eda72-171">Description</span><span class="sxs-lookup"><span data-stu-id="eda72-171">Description</span></span> | <span data-ttu-id="eda72-172">Ensemble minimal de conditions requises</span><span class="sxs-lookup"><span data-stu-id="eda72-172">Minimum mailbox requirement set version</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="eda72-173">String</span><span class="sxs-lookup"><span data-stu-id="eda72-173">String</span></span> | <span data-ttu-id="eda72-174">La date ou l’heure du rendez-vous sélectionné ou de la série a changé.</span><span class="sxs-lookup"><span data-stu-id="eda72-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="eda72-175">1.7</span><span class="sxs-lookup"><span data-stu-id="eda72-175">-17</span></span> |
|`ItemChanged`| <span data-ttu-id="eda72-176">String</span><span class="sxs-lookup"><span data-stu-id="eda72-176">String</span></span> | <span data-ttu-id="eda72-177">L’élément sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="eda72-177">The selected item has changed.</span></span> | <span data-ttu-id="eda72-178">1.5</span><span class="sxs-lookup"><span data-stu-id="eda72-178">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="eda72-179">String</span><span class="sxs-lookup"><span data-stu-id="eda72-179">String</span></span> | <span data-ttu-id="eda72-180">La liste de destinataires de l’élément sélectionné ou l’emplacement du rendez-vous a changé.</span><span class="sxs-lookup"><span data-stu-id="eda72-180">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="eda72-181">1.7</span><span class="sxs-lookup"><span data-stu-id="eda72-181">-17</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="eda72-182">String</span><span class="sxs-lookup"><span data-stu-id="eda72-182">String</span></span> | <span data-ttu-id="eda72-183">La périodicité de la série sélectionnée a changé.</span><span class="sxs-lookup"><span data-stu-id="eda72-183">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="eda72-184">1.7</span><span class="sxs-lookup"><span data-stu-id="eda72-184">-17</span></span> |

##### <a name="requirements"></a><span data-ttu-id="eda72-185">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="eda72-185">Requirements</span></span>

|<span data-ttu-id="eda72-186">Condition</span><span class="sxs-lookup"><span data-stu-id="eda72-186">Requirement</span></span>| <span data-ttu-id="eda72-187">Valeur</span><span class="sxs-lookup"><span data-stu-id="eda72-187">Value</span></span>|
|---|---|
|[<span data-ttu-id="eda72-188">Version minimale de l’ensemble des conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="eda72-188">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eda72-189">1,5</span><span class="sxs-lookup"><span data-stu-id="eda72-189">1.5</span></span> |
|[<span data-ttu-id="eda72-190">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="eda72-190">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="eda72-191">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="eda72-191">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="eda72-192">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="eda72-192">SourceProperty :String</span></span>

<span data-ttu-id="eda72-193">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="eda72-193">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="eda72-194">Type :</span><span class="sxs-lookup"><span data-stu-id="eda72-194">Type:</span></span>

*   <span data-ttu-id="eda72-195">String</span><span class="sxs-lookup"><span data-stu-id="eda72-195">String</span></span>

##### <a name="properties"></a><span data-ttu-id="eda72-196">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="eda72-196">Properties:</span></span>

|<span data-ttu-id="eda72-197">Nom</span><span class="sxs-lookup"><span data-stu-id="eda72-197">Name</span></span>| <span data-ttu-id="eda72-198">Type</span><span class="sxs-lookup"><span data-stu-id="eda72-198">Type</span></span>| <span data-ttu-id="eda72-199">Description</span><span class="sxs-lookup"><span data-stu-id="eda72-199">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="eda72-200">String</span><span class="sxs-lookup"><span data-stu-id="eda72-200">String</span></span>|<span data-ttu-id="eda72-201">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="eda72-201">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="eda72-202">String</span><span class="sxs-lookup"><span data-stu-id="eda72-202">String</span></span>|<span data-ttu-id="eda72-203">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="eda72-203">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="eda72-204">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="eda72-204">Requirements</span></span>

|<span data-ttu-id="eda72-205">Condition</span><span class="sxs-lookup"><span data-stu-id="eda72-205">Requirement</span></span>| <span data-ttu-id="eda72-206">Valeur</span><span class="sxs-lookup"><span data-stu-id="eda72-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="eda72-207">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="eda72-207">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eda72-208">1.0</span><span class="sxs-lookup"><span data-stu-id="eda72-208">1.0</span></span>|
|[<span data-ttu-id="eda72-209">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="eda72-209">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="eda72-210">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="eda72-210">Compose or read</span></span>|