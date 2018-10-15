 

# <a name="office"></a><span data-ttu-id="84e80-101">Office</span><span class="sxs-lookup"><span data-stu-id="84e80-101">Office</span></span>

<span data-ttu-id="84e80-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[Interface API partagée](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="84e80-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="84e80-104">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="84e80-104">Requirements</span></span>

|<span data-ttu-id="84e80-105">Condition requise</span><span class="sxs-lookup"><span data-stu-id="84e80-105">Requirement</span></span>| <span data-ttu-id="84e80-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="84e80-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="84e80-107">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="84e80-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="84e80-108">1.0</span><span class="sxs-lookup"><span data-stu-id="84e80-108">1.0</span></span>|
|[<span data-ttu-id="84e80-109">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="84e80-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="84e80-110">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="84e80-110">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="84e80-111">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="84e80-111">Namespaces</span></span>

<span data-ttu-id="84e80-112">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="84e80-112">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="84e80-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="84e80-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="84e80-114">Membres</span><span class="sxs-lookup"><span data-stu-id="84e80-114">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="84e80-115">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="84e80-115">AsyncResultStatus :String</span></span>

<span data-ttu-id="84e80-116">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="84e80-116">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="84e80-117">Type :</span><span class="sxs-lookup"><span data-stu-id="84e80-117">Type:</span></span>

*   <span data-ttu-id="84e80-118">Chaîne</span><span class="sxs-lookup"><span data-stu-id="84e80-118">String</span></span>

##### <a name="properties"></a><span data-ttu-id="84e80-119">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="84e80-119">Properties:</span></span>

|<span data-ttu-id="84e80-120">Nom</span><span class="sxs-lookup"><span data-stu-id="84e80-120">Name</span></span>| <span data-ttu-id="84e80-121">Type</span><span class="sxs-lookup"><span data-stu-id="84e80-121">Type</span></span>| <span data-ttu-id="84e80-122">Description</span><span class="sxs-lookup"><span data-stu-id="84e80-122">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="84e80-123">Chaîne</span><span class="sxs-lookup"><span data-stu-id="84e80-123">String</span></span>|<span data-ttu-id="84e80-124">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="84e80-124">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="84e80-125">Chaîne</span><span class="sxs-lookup"><span data-stu-id="84e80-125">String</span></span>|<span data-ttu-id="84e80-126">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="84e80-126">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="84e80-127">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="84e80-127">Requirements</span></span>

|<span data-ttu-id="84e80-128">Condition requise</span><span class="sxs-lookup"><span data-stu-id="84e80-128">Requirement</span></span>| <span data-ttu-id="84e80-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="84e80-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="84e80-130">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="84e80-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="84e80-131">1.0</span><span class="sxs-lookup"><span data-stu-id="84e80-131">1.0</span></span>|
|[<span data-ttu-id="84e80-132">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="84e80-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="84e80-133">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="84e80-133">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="84e80-134">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="84e80-134">CoercionType :String</span></span>

<span data-ttu-id="84e80-135">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="84e80-135">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="84e80-136">Type :</span><span class="sxs-lookup"><span data-stu-id="84e80-136">Type:</span></span>

*   <span data-ttu-id="84e80-137">Chaîne</span><span class="sxs-lookup"><span data-stu-id="84e80-137">String</span></span>

##### <a name="properties"></a><span data-ttu-id="84e80-138">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="84e80-138">Properties:</span></span>

|<span data-ttu-id="84e80-139">Nom</span><span class="sxs-lookup"><span data-stu-id="84e80-139">Name</span></span>| <span data-ttu-id="84e80-140">Type</span><span class="sxs-lookup"><span data-stu-id="84e80-140">Type</span></span>| <span data-ttu-id="84e80-141">Description</span><span class="sxs-lookup"><span data-stu-id="84e80-141">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="84e80-142">Chaîne</span><span class="sxs-lookup"><span data-stu-id="84e80-142">String</span></span>|<span data-ttu-id="84e80-143">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="84e80-143">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="84e80-144">Chaîne</span><span class="sxs-lookup"><span data-stu-id="84e80-144">String</span></span>|<span data-ttu-id="84e80-145">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="84e80-145">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="84e80-146">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="84e80-146">Requirements</span></span>

|<span data-ttu-id="84e80-147">Condition requise</span><span class="sxs-lookup"><span data-stu-id="84e80-147">Requirement</span></span>| <span data-ttu-id="84e80-148">Valeur</span><span class="sxs-lookup"><span data-stu-id="84e80-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="84e80-149">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="84e80-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="84e80-150">1.0</span><span class="sxs-lookup"><span data-stu-id="84e80-150">1.0</span></span>|
|[<span data-ttu-id="84e80-151">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="84e80-151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="84e80-152">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="84e80-152">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="84e80-153">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="84e80-153">SourceProperty :String</span></span>

<span data-ttu-id="84e80-154">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="84e80-154">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="84e80-155">Type :</span><span class="sxs-lookup"><span data-stu-id="84e80-155">Type:</span></span>

*   <span data-ttu-id="84e80-156">Chaîne</span><span class="sxs-lookup"><span data-stu-id="84e80-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="84e80-157">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="84e80-157">Properties:</span></span>

|<span data-ttu-id="84e80-158">Nom</span><span class="sxs-lookup"><span data-stu-id="84e80-158">Name</span></span>| <span data-ttu-id="84e80-159">Type</span><span class="sxs-lookup"><span data-stu-id="84e80-159">Type</span></span>| <span data-ttu-id="84e80-160">Description</span><span class="sxs-lookup"><span data-stu-id="84e80-160">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="84e80-161">Chaîne</span><span class="sxs-lookup"><span data-stu-id="84e80-161">String</span></span>|<span data-ttu-id="84e80-162">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="84e80-162">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="84e80-163">Chaîne</span><span class="sxs-lookup"><span data-stu-id="84e80-163">String</span></span>|<span data-ttu-id="84e80-164">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="84e80-164">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="84e80-165">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="84e80-165">Requirements</span></span>

|<span data-ttu-id="84e80-166">Condition requise</span><span class="sxs-lookup"><span data-stu-id="84e80-166">Requirement</span></span>| <span data-ttu-id="84e80-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="84e80-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="84e80-168">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="84e80-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="84e80-169">1.0</span><span class="sxs-lookup"><span data-stu-id="84e80-169">1.0</span></span>|
|[<span data-ttu-id="84e80-170">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="84e80-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="84e80-171">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="84e80-171">Compose or read</span></span>|