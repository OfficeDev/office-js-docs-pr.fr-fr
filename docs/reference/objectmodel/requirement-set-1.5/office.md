# <a name="office"></a><span data-ttu-id="d65b3-101">Office</span><span class="sxs-lookup"><span data-stu-id="d65b3-101">Office</span></span>

<span data-ttu-id="d65b3-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[Interface API partagée](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="d65b3-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="d65b3-104">Configurations requises</span><span class="sxs-lookup"><span data-stu-id="d65b3-104">Requirements</span></span>

|<span data-ttu-id="d65b3-105">Condition</span><span class="sxs-lookup"><span data-stu-id="d65b3-105">Requirement</span></span>| <span data-ttu-id="d65b3-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="d65b3-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="d65b3-107">Version minimale de l’ensemble de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d65b3-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d65b3-108">1.0</span><span class="sxs-lookup"><span data-stu-id="d65b3-108">1.0</span></span>|
|[<span data-ttu-id="d65b3-109">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d65b3-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d65b3-110">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="d65b3-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="d65b3-111">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="d65b3-111">Members and methods</span></span>

| <span data-ttu-id="d65b3-112">Membre</span><span class="sxs-lookup"><span data-stu-id="d65b3-112">Member</span></span> | <span data-ttu-id="d65b3-113">Type</span><span class="sxs-lookup"><span data-stu-id="d65b3-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="d65b3-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="d65b3-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="d65b3-115">Membre</span><span class="sxs-lookup"><span data-stu-id="d65b3-115">Member</span></span> |
| [<span data-ttu-id="d65b3-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="d65b3-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="d65b3-117">Membre</span><span class="sxs-lookup"><span data-stu-id="d65b3-117">Member</span></span> |
| [<span data-ttu-id="d65b3-118">EventType</span><span class="sxs-lookup"><span data-stu-id="d65b3-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="d65b3-119">Membre</span><span class="sxs-lookup"><span data-stu-id="d65b3-119">Member</span></span> |
| [<span data-ttu-id="d65b3-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="d65b3-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="d65b3-121">Membre</span><span class="sxs-lookup"><span data-stu-id="d65b3-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="d65b3-122">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="d65b3-122">Namespaces</span></span>

<span data-ttu-id="d65b3-123">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="d65b3-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="d65b3-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="d65b3-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="d65b3-125">Membres</span><span class="sxs-lookup"><span data-stu-id="d65b3-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="d65b3-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="d65b3-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="d65b3-127">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="d65b3-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="d65b3-128">Type :</span><span class="sxs-lookup"><span data-stu-id="d65b3-128">Type:</span></span>

*   <span data-ttu-id="d65b3-129">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d65b3-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d65b3-130">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="d65b3-130">Properties:</span></span>

|<span data-ttu-id="d65b3-131">Nom</span><span class="sxs-lookup"><span data-stu-id="d65b3-131">Name</span></span>| <span data-ttu-id="d65b3-132">Type</span><span class="sxs-lookup"><span data-stu-id="d65b3-132">Type</span></span>| <span data-ttu-id="d65b3-133">Description</span><span class="sxs-lookup"><span data-stu-id="d65b3-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="d65b3-134">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d65b3-134">String</span></span>|<span data-ttu-id="d65b3-135">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="d65b3-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="d65b3-136">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d65b3-136">String</span></span>|<span data-ttu-id="d65b3-137">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="d65b3-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d65b3-138">Configurations requises</span><span class="sxs-lookup"><span data-stu-id="d65b3-138">Requirements</span></span>

|<span data-ttu-id="d65b3-139">Condition</span><span class="sxs-lookup"><span data-stu-id="d65b3-139">Requirement</span></span>| <span data-ttu-id="d65b3-140">Valeur</span><span class="sxs-lookup"><span data-stu-id="d65b3-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="d65b3-141">Version minimale de l’ensemble de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d65b3-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d65b3-142">1.0</span><span class="sxs-lookup"><span data-stu-id="d65b3-142">1.0</span></span>|
|[<span data-ttu-id="d65b3-143">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d65b3-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d65b3-144">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="d65b3-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="d65b3-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="d65b3-145">CoercionType :String</span></span>

<span data-ttu-id="d65b3-146">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="d65b3-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d65b3-147">Type :</span><span class="sxs-lookup"><span data-stu-id="d65b3-147">Type:</span></span>

*   <span data-ttu-id="d65b3-148">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d65b3-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d65b3-149">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="d65b3-149">Properties:</span></span>

|<span data-ttu-id="d65b3-150">Nom</span><span class="sxs-lookup"><span data-stu-id="d65b3-150">Name</span></span>| <span data-ttu-id="d65b3-151">Type</span><span class="sxs-lookup"><span data-stu-id="d65b3-151">Type</span></span>| <span data-ttu-id="d65b3-152">Description</span><span class="sxs-lookup"><span data-stu-id="d65b3-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="d65b3-153">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d65b3-153">String</span></span>|<span data-ttu-id="d65b3-154">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="d65b3-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="d65b3-155">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d65b3-155">String</span></span>|<span data-ttu-id="d65b3-156">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="d65b3-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d65b3-157">Configurations requises</span><span class="sxs-lookup"><span data-stu-id="d65b3-157">Requirements</span></span>

|<span data-ttu-id="d65b3-158">Condition</span><span class="sxs-lookup"><span data-stu-id="d65b3-158">Requirement</span></span>| <span data-ttu-id="d65b3-159">Valeur</span><span class="sxs-lookup"><span data-stu-id="d65b3-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="d65b3-160">Version minimale de l’ensemble de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d65b3-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d65b3-161">1.0</span><span class="sxs-lookup"><span data-stu-id="d65b3-161">1.0</span></span>|
|[<span data-ttu-id="d65b3-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d65b3-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d65b3-163">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="d65b3-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="d65b3-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="d65b3-164">EventType :String</span></span>

<span data-ttu-id="d65b3-165">Spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="d65b3-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="d65b3-166">Type :</span><span class="sxs-lookup"><span data-stu-id="d65b3-166">Type:</span></span>

*   <span data-ttu-id="d65b3-167">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d65b3-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d65b3-168">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="d65b3-168">Properties:</span></span>

| <span data-ttu-id="d65b3-169">Nom</span><span class="sxs-lookup"><span data-stu-id="d65b3-169">Name</span></span> | <span data-ttu-id="d65b3-170">Type</span><span class="sxs-lookup"><span data-stu-id="d65b3-170">Type</span></span> | <span data-ttu-id="d65b3-171">Description</span><span class="sxs-lookup"><span data-stu-id="d65b3-171">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="d65b3-172">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d65b3-172">String</span></span> | <span data-ttu-id="d65b3-173">L’élément sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="d65b3-173">The selected item has changed.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d65b3-174">Configurations requises</span><span class="sxs-lookup"><span data-stu-id="d65b3-174">Requirements</span></span>

|<span data-ttu-id="d65b3-175">Condition</span><span class="sxs-lookup"><span data-stu-id="d65b3-175">Requirement</span></span>| <span data-ttu-id="d65b3-176">Valeur</span><span class="sxs-lookup"><span data-stu-id="d65b3-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="d65b3-177">Version minimale de l’ensemble de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d65b3-177">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d65b3-178">1,5</span><span class="sxs-lookup"><span data-stu-id="d65b3-178">1.5</span></span> |
|[<span data-ttu-id="d65b3-179">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d65b3-179">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d65b3-180">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="d65b3-180">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="d65b3-181">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="d65b3-181">SourceProperty :String</span></span>

<span data-ttu-id="d65b3-182">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="d65b3-182">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d65b3-183">Type :</span><span class="sxs-lookup"><span data-stu-id="d65b3-183">Type:</span></span>

*   <span data-ttu-id="d65b3-184">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d65b3-184">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d65b3-185">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="d65b3-185">Properties:</span></span>

|<span data-ttu-id="d65b3-186">Nom</span><span class="sxs-lookup"><span data-stu-id="d65b3-186">Name</span></span>| <span data-ttu-id="d65b3-187">Type</span><span class="sxs-lookup"><span data-stu-id="d65b3-187">Type</span></span>| <span data-ttu-id="d65b3-188">Description</span><span class="sxs-lookup"><span data-stu-id="d65b3-188">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="d65b3-189">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d65b3-189">String</span></span>|<span data-ttu-id="d65b3-190">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="d65b3-190">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="d65b3-191">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d65b3-191">String</span></span>|<span data-ttu-id="d65b3-192">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="d65b3-192">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d65b3-193">Configurations requises</span><span class="sxs-lookup"><span data-stu-id="d65b3-193">Requirements</span></span>

|<span data-ttu-id="d65b3-194">Condition</span><span class="sxs-lookup"><span data-stu-id="d65b3-194">Requirement</span></span>| <span data-ttu-id="d65b3-195">Valeur</span><span class="sxs-lookup"><span data-stu-id="d65b3-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="d65b3-196">Version minimale de l’ensemble de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d65b3-196">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d65b3-197">1.0</span><span class="sxs-lookup"><span data-stu-id="d65b3-197">1.0</span></span>|
|[<span data-ttu-id="d65b3-198">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d65b3-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d65b3-199">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="d65b3-199">Compose or read</span></span>|