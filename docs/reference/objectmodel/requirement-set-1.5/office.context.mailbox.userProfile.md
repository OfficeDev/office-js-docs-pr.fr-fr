# <a name="userprofile"></a><span data-ttu-id="d5f89-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="d5f89-101">userProfile</span></span>

### <span data-ttu-id="d5f89-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="d5f89-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="d5f89-104">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d5f89-104">Requirements</span></span>

|<span data-ttu-id="d5f89-105">Condition requise</span><span class="sxs-lookup"><span data-stu-id="d5f89-105">Requirement</span></span>| <span data-ttu-id="d5f89-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="d5f89-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="d5f89-107">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d5f89-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d5f89-108">1.0</span><span class="sxs-lookup"><span data-stu-id="d5f89-108">1.0</span></span>|
|[<span data-ttu-id="d5f89-109">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="d5f89-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d5f89-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d5f89-110">ReadItem</span></span>|
|[<span data-ttu-id="d5f89-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d5f89-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d5f89-112">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="d5f89-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="d5f89-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="d5f89-113">Members and methods</span></span>

| <span data-ttu-id="d5f89-114">Membre</span><span class="sxs-lookup"><span data-stu-id="d5f89-114">Member</span></span> | <span data-ttu-id="d5f89-115">Type</span><span class="sxs-lookup"><span data-stu-id="d5f89-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="d5f89-116">displayName</span><span class="sxs-lookup"><span data-stu-id="d5f89-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="d5f89-117">Membre</span><span class="sxs-lookup"><span data-stu-id="d5f89-117">Member</span></span> |
| [<span data-ttu-id="d5f89-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="d5f89-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="d5f89-119">Membre</span><span class="sxs-lookup"><span data-stu-id="d5f89-119">Member</span></span> |
| [<span data-ttu-id="d5f89-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="d5f89-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="d5f89-121">Membre</span><span class="sxs-lookup"><span data-stu-id="d5f89-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="d5f89-122">Membres</span><span class="sxs-lookup"><span data-stu-id="d5f89-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="d5f89-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="d5f89-123">displayName :String</span></span>

<span data-ttu-id="d5f89-124">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="d5f89-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="d5f89-125">Type :</span><span class="sxs-lookup"><span data-stu-id="d5f89-125">Type:</span></span>

*   <span data-ttu-id="d5f89-126">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d5f89-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d5f89-127">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d5f89-127">Requirements</span></span>

|<span data-ttu-id="d5f89-128">Condition requise</span><span class="sxs-lookup"><span data-stu-id="d5f89-128">Requirement</span></span>| <span data-ttu-id="d5f89-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="d5f89-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="d5f89-130">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d5f89-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d5f89-131">1.0</span><span class="sxs-lookup"><span data-stu-id="d5f89-131">1.0</span></span>|
|[<span data-ttu-id="d5f89-132">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="d5f89-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d5f89-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d5f89-133">ReadItem</span></span>|
|[<span data-ttu-id="d5f89-134">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d5f89-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d5f89-135">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="d5f89-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d5f89-136">Exemple</span><span class="sxs-lookup"><span data-stu-id="d5f89-136">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="d5f89-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="d5f89-137">emailAddress :String</span></span>

<span data-ttu-id="d5f89-138">Obtient l’adresse e-mail SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="d5f89-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="d5f89-139">Type :</span><span class="sxs-lookup"><span data-stu-id="d5f89-139">Type:</span></span>

*   <span data-ttu-id="d5f89-140">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d5f89-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d5f89-141">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d5f89-141">Requirements</span></span>

|<span data-ttu-id="d5f89-142">Condition requise</span><span class="sxs-lookup"><span data-stu-id="d5f89-142">Requirement</span></span>| <span data-ttu-id="d5f89-143">Valeur</span><span class="sxs-lookup"><span data-stu-id="d5f89-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="d5f89-144">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d5f89-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d5f89-145">1.0</span><span class="sxs-lookup"><span data-stu-id="d5f89-145">1.0</span></span>|
|[<span data-ttu-id="d5f89-146">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="d5f89-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d5f89-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d5f89-147">ReadItem</span></span>|
|[<span data-ttu-id="d5f89-148">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d5f89-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d5f89-149">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="d5f89-149">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d5f89-150">Exemple</span><span class="sxs-lookup"><span data-stu-id="d5f89-150">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="d5f89-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="d5f89-151">timeZone :String</span></span>

<span data-ttu-id="d5f89-152">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="d5f89-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="d5f89-153">Type :</span><span class="sxs-lookup"><span data-stu-id="d5f89-153">Type:</span></span>

*   <span data-ttu-id="d5f89-154">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d5f89-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d5f89-155">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d5f89-155">Requirements</span></span>

|<span data-ttu-id="d5f89-156">Condition requise</span><span class="sxs-lookup"><span data-stu-id="d5f89-156">Requirement</span></span>| <span data-ttu-id="d5f89-157">Valeur</span><span class="sxs-lookup"><span data-stu-id="d5f89-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="d5f89-158">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d5f89-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d5f89-159">1.0</span><span class="sxs-lookup"><span data-stu-id="d5f89-159">1.0</span></span>|
|[<span data-ttu-id="d5f89-160">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="d5f89-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d5f89-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d5f89-161">ReadItem</span></span>|
|[<span data-ttu-id="d5f89-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d5f89-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d5f89-163">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="d5f89-163">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d5f89-164">Exemple</span><span class="sxs-lookup"><span data-stu-id="d5f89-164">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```