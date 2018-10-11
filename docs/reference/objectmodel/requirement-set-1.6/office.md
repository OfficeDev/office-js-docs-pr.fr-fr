 

# <a name="office"></a><span data-ttu-id="88df6-101">Office</span><span class="sxs-lookup"><span data-stu-id="88df6-101">Office</span></span>

<span data-ttu-id="88df6-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[Interface API partagée](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="88df6-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="88df6-104">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="88df6-104">Requirements</span></span>

|<span data-ttu-id="88df6-105">Condition requise</span><span class="sxs-lookup"><span data-stu-id="88df6-105">Requirement</span></span>| <span data-ttu-id="88df6-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="88df6-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="88df6-107">Version minimale de l’ensemble de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="88df6-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88df6-108">1.0</span><span class="sxs-lookup"><span data-stu-id="88df6-108">1.0</span></span>|
|[<span data-ttu-id="88df6-109">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="88df6-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88df6-110">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="88df6-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="88df6-111">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="88df6-111">Members and methods</span></span>

| <span data-ttu-id="88df6-112">Membre</span><span class="sxs-lookup"><span data-stu-id="88df6-112">Member</span></span> | <span data-ttu-id="88df6-113">Type</span><span class="sxs-lookup"><span data-stu-id="88df6-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="88df6-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="88df6-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="88df6-115">Membre</span><span class="sxs-lookup"><span data-stu-id="88df6-115">Member</span></span> |
| [<span data-ttu-id="88df6-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="88df6-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="88df6-117">Membre</span><span class="sxs-lookup"><span data-stu-id="88df6-117">Member</span></span> |
| [<span data-ttu-id="88df6-118">EventType</span><span class="sxs-lookup"><span data-stu-id="88df6-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="88df6-119">Membre</span><span class="sxs-lookup"><span data-stu-id="88df6-119">Member</span></span> |
| [<span data-ttu-id="88df6-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="88df6-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="88df6-121">Membre</span><span class="sxs-lookup"><span data-stu-id="88df6-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="88df6-122">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="88df6-122">Namespaces</span></span>

<span data-ttu-id="88df6-123">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="88df6-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="88df6-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="88df6-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="88df6-125">Membres</span><span class="sxs-lookup"><span data-stu-id="88df6-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="88df6-126">AsyncResultStatus : Chaîne</span><span class="sxs-lookup"><span data-stu-id="88df6-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="88df6-127">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="88df6-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="88df6-128">Type :</span><span class="sxs-lookup"><span data-stu-id="88df6-128">Type:</span></span>

*   <span data-ttu-id="88df6-129">String</span><span class="sxs-lookup"><span data-stu-id="88df6-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="88df6-130">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="88df6-130">Properties:</span></span>

|<span data-ttu-id="88df6-131">Nom</span><span class="sxs-lookup"><span data-stu-id="88df6-131">Name</span></span>| <span data-ttu-id="88df6-132">Type</span><span class="sxs-lookup"><span data-stu-id="88df6-132">Type</span></span>| <span data-ttu-id="88df6-133">Description</span><span class="sxs-lookup"><span data-stu-id="88df6-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="88df6-134">String</span><span class="sxs-lookup"><span data-stu-id="88df6-134">String</span></span>|<span data-ttu-id="88df6-135">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="88df6-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="88df6-136">String</span><span class="sxs-lookup"><span data-stu-id="88df6-136">String</span></span>|<span data-ttu-id="88df6-137">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="88df6-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="88df6-138">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="88df6-138">Requirements</span></span>

|<span data-ttu-id="88df6-139">Condition requise</span><span class="sxs-lookup"><span data-stu-id="88df6-139">Requirement</span></span>| <span data-ttu-id="88df6-140">Valeur</span><span class="sxs-lookup"><span data-stu-id="88df6-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="88df6-141">Version minimale de l’ensemble de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="88df6-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88df6-142">1.0</span><span class="sxs-lookup"><span data-stu-id="88df6-142">1.0</span></span>|
|[<span data-ttu-id="88df6-143">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="88df6-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88df6-144">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="88df6-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="88df6-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="88df6-145">CoercionType :String</span></span>

<span data-ttu-id="88df6-146">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="88df6-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="88df6-147">Type :</span><span class="sxs-lookup"><span data-stu-id="88df6-147">Type:</span></span>

*   <span data-ttu-id="88df6-148">String</span><span class="sxs-lookup"><span data-stu-id="88df6-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="88df6-149">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="88df6-149">Properties:</span></span>

|<span data-ttu-id="88df6-150">Nom</span><span class="sxs-lookup"><span data-stu-id="88df6-150">Name</span></span>| <span data-ttu-id="88df6-151">Type</span><span class="sxs-lookup"><span data-stu-id="88df6-151">Type</span></span>| <span data-ttu-id="88df6-152">Description</span><span class="sxs-lookup"><span data-stu-id="88df6-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="88df6-153">String</span><span class="sxs-lookup"><span data-stu-id="88df6-153">String</span></span>|<span data-ttu-id="88df6-154">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="88df6-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="88df6-155">String</span><span class="sxs-lookup"><span data-stu-id="88df6-155">String</span></span>|<span data-ttu-id="88df6-156">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="88df6-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="88df6-157">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="88df6-157">Requirements</span></span>

|<span data-ttu-id="88df6-158">Condition requise</span><span class="sxs-lookup"><span data-stu-id="88df6-158">Requirement</span></span>| <span data-ttu-id="88df6-159">Valeur</span><span class="sxs-lookup"><span data-stu-id="88df6-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="88df6-160">Version minimale de l’ensemble de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="88df6-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88df6-161">1.0</span><span class="sxs-lookup"><span data-stu-id="88df6-161">1.0</span></span>|
|[<span data-ttu-id="88df6-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="88df6-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88df6-163">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="88df6-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="88df6-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="88df6-164">EventType :String</span></span>

<span data-ttu-id="88df6-165">Spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="88df6-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="88df6-166">Type :</span><span class="sxs-lookup"><span data-stu-id="88df6-166">Type:</span></span>

*   <span data-ttu-id="88df6-167">String</span><span class="sxs-lookup"><span data-stu-id="88df6-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="88df6-168">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="88df6-168">Properties:</span></span>

| <span data-ttu-id="88df6-169">Nom</span><span class="sxs-lookup"><span data-stu-id="88df6-169">Name</span></span> | <span data-ttu-id="88df6-170">Type</span><span class="sxs-lookup"><span data-stu-id="88df6-170">Type</span></span> | <span data-ttu-id="88df6-171">Description</span><span class="sxs-lookup"><span data-stu-id="88df6-171">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="88df6-172">String</span><span class="sxs-lookup"><span data-stu-id="88df6-172">String</span></span> | <span data-ttu-id="88df6-173">L’élément sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="88df6-173">The selected item has changed.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="88df6-174">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="88df6-174">Requirements</span></span>

|<span data-ttu-id="88df6-175">Condition requise</span><span class="sxs-lookup"><span data-stu-id="88df6-175">Requirement</span></span>| <span data-ttu-id="88df6-176">Valeur</span><span class="sxs-lookup"><span data-stu-id="88df6-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="88df6-177">Version minimale de l’ensemble de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="88df6-177">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88df6-178">1,5</span><span class="sxs-lookup"><span data-stu-id="88df6-178">1.5</span></span> |
|[<span data-ttu-id="88df6-179">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="88df6-179">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88df6-180">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="88df6-180">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="88df6-181">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="88df6-181">SourceProperty :String</span></span>

<span data-ttu-id="88df6-182">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="88df6-182">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="88df6-183">Type :</span><span class="sxs-lookup"><span data-stu-id="88df6-183">Type:</span></span>

*   <span data-ttu-id="88df6-184">String</span><span class="sxs-lookup"><span data-stu-id="88df6-184">String</span></span>

##### <a name="properties"></a><span data-ttu-id="88df6-185">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="88df6-185">Properties:</span></span>

|<span data-ttu-id="88df6-186">Nom</span><span class="sxs-lookup"><span data-stu-id="88df6-186">Name</span></span>| <span data-ttu-id="88df6-187">Type</span><span class="sxs-lookup"><span data-stu-id="88df6-187">Type</span></span>| <span data-ttu-id="88df6-188">Description</span><span class="sxs-lookup"><span data-stu-id="88df6-188">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="88df6-189">String</span><span class="sxs-lookup"><span data-stu-id="88df6-189">String</span></span>|<span data-ttu-id="88df6-190">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="88df6-190">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="88df6-191">String</span><span class="sxs-lookup"><span data-stu-id="88df6-191">String</span></span>|<span data-ttu-id="88df6-192">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="88df6-192">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="88df6-193">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="88df6-193">Requirements</span></span>

|<span data-ttu-id="88df6-194">Condition requise</span><span class="sxs-lookup"><span data-stu-id="88df6-194">Requirement</span></span>| <span data-ttu-id="88df6-195">Valeur</span><span class="sxs-lookup"><span data-stu-id="88df6-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="88df6-196">Version minimale de l’ensemble de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="88df6-196">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88df6-197">1.0</span><span class="sxs-lookup"><span data-stu-id="88df6-197">1.0</span></span>|
|[<span data-ttu-id="88df6-198">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="88df6-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="88df6-199">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="88df6-199">Compose or read</span></span>|