# <a name="office"></a><span data-ttu-id="a8ac0-101">Office</span><span class="sxs-lookup"><span data-stu-id="a8ac0-101">Office</span></span>

<span data-ttu-id="a8ac0-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[Interface API partagée](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="a8ac0-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="a8ac0-104">Configurations requises</span><span class="sxs-lookup"><span data-stu-id="a8ac0-104">Requirements</span></span>

|<span data-ttu-id="a8ac0-105">Condition requise</span><span class="sxs-lookup"><span data-stu-id="a8ac0-105">Requirement</span></span>| <span data-ttu-id="a8ac0-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="a8ac0-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="a8ac0-107">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a8ac0-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a8ac0-108">1.0</span><span class="sxs-lookup"><span data-stu-id="a8ac0-108">1.0</span></span>|
|[<span data-ttu-id="a8ac0-109">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a8ac0-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a8ac0-110">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="a8ac0-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="a8ac0-111">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="a8ac0-111">Members and methods</span></span>

| <span data-ttu-id="a8ac0-112">Membre</span><span class="sxs-lookup"><span data-stu-id="a8ac0-112">Member</span></span> | <span data-ttu-id="a8ac0-113">Type</span><span class="sxs-lookup"><span data-stu-id="a8ac0-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="a8ac0-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="a8ac0-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="a8ac0-115">Membre</span><span class="sxs-lookup"><span data-stu-id="a8ac0-115">Member</span></span> |
| [<span data-ttu-id="a8ac0-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="a8ac0-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="a8ac0-117">Membre</span><span class="sxs-lookup"><span data-stu-id="a8ac0-117">Member</span></span> |
| [<span data-ttu-id="a8ac0-118">EventType</span><span class="sxs-lookup"><span data-stu-id="a8ac0-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="a8ac0-119">Membre</span><span class="sxs-lookup"><span data-stu-id="a8ac0-119">Member</span></span> |
| [<span data-ttu-id="a8ac0-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="a8ac0-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="a8ac0-121">Membre</span><span class="sxs-lookup"><span data-stu-id="a8ac0-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="a8ac0-122">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="a8ac0-122">Namespaces</span></span>

<span data-ttu-id="a8ac0-123">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="a8ac0-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="a8ac0-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="a8ac0-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="a8ac0-125">Membres</span><span class="sxs-lookup"><span data-stu-id="a8ac0-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="a8ac0-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="a8ac0-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="a8ac0-127">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="a8ac0-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="a8ac0-128">Type :</span><span class="sxs-lookup"><span data-stu-id="a8ac0-128">Type:</span></span>

*   <span data-ttu-id="a8ac0-129">String</span><span class="sxs-lookup"><span data-stu-id="a8ac0-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a8ac0-130">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="a8ac0-130">Properties:</span></span>

|<span data-ttu-id="a8ac0-131">Nom</span><span class="sxs-lookup"><span data-stu-id="a8ac0-131">Name</span></span>| <span data-ttu-id="a8ac0-132">Type</span><span class="sxs-lookup"><span data-stu-id="a8ac0-132">Type</span></span>| <span data-ttu-id="a8ac0-133">Description</span><span class="sxs-lookup"><span data-stu-id="a8ac0-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="a8ac0-134">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a8ac0-134">String</span></span>|<span data-ttu-id="a8ac0-135">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="a8ac0-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="a8ac0-136">String</span><span class="sxs-lookup"><span data-stu-id="a8ac0-136">String</span></span>|<span data-ttu-id="a8ac0-137">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="a8ac0-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a8ac0-138">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a8ac0-138">Requirements</span></span>

|<span data-ttu-id="a8ac0-139">Condition requise</span><span class="sxs-lookup"><span data-stu-id="a8ac0-139">Requirement</span></span>| <span data-ttu-id="a8ac0-140">Valeur</span><span class="sxs-lookup"><span data-stu-id="a8ac0-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="a8ac0-141">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a8ac0-141">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a8ac0-142">1.0</span><span class="sxs-lookup"><span data-stu-id="a8ac0-142">1.0</span></span>|
|[<span data-ttu-id="a8ac0-143">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a8ac0-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a8ac0-144">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="a8ac0-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="a8ac0-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="a8ac0-145">CoercionType :String</span></span>

<span data-ttu-id="a8ac0-146">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="a8ac0-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a8ac0-147">Type :</span><span class="sxs-lookup"><span data-stu-id="a8ac0-147">Type:</span></span>

*   <span data-ttu-id="a8ac0-148">String</span><span class="sxs-lookup"><span data-stu-id="a8ac0-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a8ac0-149">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="a8ac0-149">Properties:</span></span>

|<span data-ttu-id="a8ac0-150">Nom</span><span class="sxs-lookup"><span data-stu-id="a8ac0-150">Name</span></span>| <span data-ttu-id="a8ac0-151">Type</span><span class="sxs-lookup"><span data-stu-id="a8ac0-151">Type</span></span>| <span data-ttu-id="a8ac0-152">Description</span><span class="sxs-lookup"><span data-stu-id="a8ac0-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="a8ac0-153">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a8ac0-153">String</span></span>|<span data-ttu-id="a8ac0-154">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="a8ac0-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="a8ac0-155">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a8ac0-155">String</span></span>|<span data-ttu-id="a8ac0-156">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="a8ac0-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a8ac0-157">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a8ac0-157">Requirements</span></span>

|<span data-ttu-id="a8ac0-158">Condition requise</span><span class="sxs-lookup"><span data-stu-id="a8ac0-158">Requirement</span></span>| <span data-ttu-id="a8ac0-159">Valeur</span><span class="sxs-lookup"><span data-stu-id="a8ac0-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="a8ac0-160">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a8ac0-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a8ac0-161">1.0</span><span class="sxs-lookup"><span data-stu-id="a8ac0-161">1.0</span></span>|
|[<span data-ttu-id="a8ac0-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a8ac0-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a8ac0-163">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="a8ac0-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="a8ac0-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="a8ac0-164">EventType :String</span></span>

<span data-ttu-id="a8ac0-165">Spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="a8ac0-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="a8ac0-166">Type :</span><span class="sxs-lookup"><span data-stu-id="a8ac0-166">Type:</span></span>

*   <span data-ttu-id="a8ac0-167">String</span><span class="sxs-lookup"><span data-stu-id="a8ac0-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a8ac0-168">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="a8ac0-168">Properties:</span></span>

| <span data-ttu-id="a8ac0-169">Nom</span><span class="sxs-lookup"><span data-stu-id="a8ac0-169">Name</span></span> | <span data-ttu-id="a8ac0-170">Type</span><span class="sxs-lookup"><span data-stu-id="a8ac0-170">Type</span></span> | <span data-ttu-id="a8ac0-171">Description</span><span class="sxs-lookup"><span data-stu-id="a8ac0-171">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="a8ac0-172">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a8ac0-172">String</span></span> | <span data-ttu-id="a8ac0-173">L’élément sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="a8ac0-173">The selected item has changed.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a8ac0-174">Configurations requises</span><span class="sxs-lookup"><span data-stu-id="a8ac0-174">Requirements</span></span>

|<span data-ttu-id="a8ac0-175">Condition requise</span><span class="sxs-lookup"><span data-stu-id="a8ac0-175">Requirement</span></span>| <span data-ttu-id="a8ac0-176">Valeur</span><span class="sxs-lookup"><span data-stu-id="a8ac0-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="a8ac0-177">Version minimale de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a8ac0-177">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a8ac0-178">1,5</span><span class="sxs-lookup"><span data-stu-id="a8ac0-178">1.5</span></span> |
|[<span data-ttu-id="a8ac0-179">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a8ac0-179">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a8ac0-180">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="a8ac0-180">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="a8ac0-181">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="a8ac0-181">SourceProperty :String</span></span>

<span data-ttu-id="a8ac0-182">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="a8ac0-182">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a8ac0-183">Type :</span><span class="sxs-lookup"><span data-stu-id="a8ac0-183">Type:</span></span>

*   <span data-ttu-id="a8ac0-184">String</span><span class="sxs-lookup"><span data-stu-id="a8ac0-184">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a8ac0-185">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="a8ac0-185">Properties:</span></span>

|<span data-ttu-id="a8ac0-186">Nom</span><span class="sxs-lookup"><span data-stu-id="a8ac0-186">Name</span></span>| <span data-ttu-id="a8ac0-187">Type</span><span class="sxs-lookup"><span data-stu-id="a8ac0-187">Type</span></span>| <span data-ttu-id="a8ac0-188">Description</span><span class="sxs-lookup"><span data-stu-id="a8ac0-188">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="a8ac0-189">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a8ac0-189">String</span></span>|<span data-ttu-id="a8ac0-190">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="a8ac0-190">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="a8ac0-191">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a8ac0-191">String</span></span>|<span data-ttu-id="a8ac0-192">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="a8ac0-192">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a8ac0-193">Configurations requises</span><span class="sxs-lookup"><span data-stu-id="a8ac0-193">Requirements</span></span>

|<span data-ttu-id="a8ac0-194">Condition requise</span><span class="sxs-lookup"><span data-stu-id="a8ac0-194">Requirement</span></span>| <span data-ttu-id="a8ac0-195">Valeur</span><span class="sxs-lookup"><span data-stu-id="a8ac0-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="a8ac0-196">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a8ac0-196">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a8ac0-197">1.0</span><span class="sxs-lookup"><span data-stu-id="a8ac0-197">1.0</span></span>|
|[<span data-ttu-id="a8ac0-198">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a8ac0-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a8ac0-199">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="a8ac0-199">Compose or read</span></span>|