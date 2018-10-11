# <a name="userprofile"></a><span data-ttu-id="c1e6a-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="c1e6a-101">userProfile</span></span>

### <span data-ttu-id="c1e6a-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="c1e6a-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="c1e6a-104">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c1e6a-104">Requirements</span></span>

|<span data-ttu-id="c1e6a-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c1e6a-105">Requirement</span></span>| <span data-ttu-id="c1e6a-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="c1e6a-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="c1e6a-107">Version minimale de l’ensemble de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c1e6a-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c1e6a-108">1.0</span><span class="sxs-lookup"><span data-stu-id="c1e6a-108">1.0</span></span>|
|[<span data-ttu-id="c1e6a-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c1e6a-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c1e6a-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c1e6a-110">ReadItem</span></span>|
|[<span data-ttu-id="c1e6a-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c1e6a-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c1e6a-112">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c1e6a-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c1e6a-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="c1e6a-113">Members and methods</span></span>

| <span data-ttu-id="c1e6a-114">Membre</span><span class="sxs-lookup"><span data-stu-id="c1e6a-114">Member</span></span> | <span data-ttu-id="c1e6a-115">Type</span><span class="sxs-lookup"><span data-stu-id="c1e6a-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c1e6a-116">displayName</span><span class="sxs-lookup"><span data-stu-id="c1e6a-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="c1e6a-117">Membre</span><span class="sxs-lookup"><span data-stu-id="c1e6a-117">Member</span></span> |
| [<span data-ttu-id="c1e6a-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="c1e6a-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="c1e6a-119">Membre</span><span class="sxs-lookup"><span data-stu-id="c1e6a-119">Member</span></span> |
| [<span data-ttu-id="c1e6a-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="c1e6a-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="c1e6a-121">Membre</span><span class="sxs-lookup"><span data-stu-id="c1e6a-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="c1e6a-122">Membres</span><span class="sxs-lookup"><span data-stu-id="c1e6a-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="c1e6a-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="c1e6a-123">displayName :String</span></span>

<span data-ttu-id="c1e6a-124">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="c1e6a-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="c1e6a-125">Type :</span><span class="sxs-lookup"><span data-stu-id="c1e6a-125">Type:</span></span>

*   <span data-ttu-id="c1e6a-126">String</span><span class="sxs-lookup"><span data-stu-id="c1e6a-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c1e6a-127">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c1e6a-127">Requirements</span></span>

|<span data-ttu-id="c1e6a-128">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c1e6a-128">Requirement</span></span>| <span data-ttu-id="c1e6a-129">Valeur</span><span class="sxs-lookup"><span data-stu-id="c1e6a-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="c1e6a-130">Version minimale de l’ensemble de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c1e6a-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c1e6a-131">1.0</span><span class="sxs-lookup"><span data-stu-id="c1e6a-131">1.0</span></span>|
|[<span data-ttu-id="c1e6a-132">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c1e6a-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c1e6a-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c1e6a-133">ReadItem</span></span>|
|[<span data-ttu-id="c1e6a-134">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c1e6a-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c1e6a-135">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c1e6a-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c1e6a-136">Exemple</span><span class="sxs-lookup"><span data-stu-id="c1e6a-136">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="c1e6a-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="c1e6a-137">emailAddress :String</span></span>

<span data-ttu-id="c1e6a-138">Obtient l’adresse e-mail SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="c1e6a-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="c1e6a-139">Type :</span><span class="sxs-lookup"><span data-stu-id="c1e6a-139">Type:</span></span>

*   <span data-ttu-id="c1e6a-140">String</span><span class="sxs-lookup"><span data-stu-id="c1e6a-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c1e6a-141">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c1e6a-141">Requirements</span></span>

|<span data-ttu-id="c1e6a-142">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c1e6a-142">Requirement</span></span>| <span data-ttu-id="c1e6a-143">Valeur</span><span class="sxs-lookup"><span data-stu-id="c1e6a-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="c1e6a-144">Version minimale de l’ensemble de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c1e6a-144">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c1e6a-145">1.0</span><span class="sxs-lookup"><span data-stu-id="c1e6a-145">1.0</span></span>|
|[<span data-ttu-id="c1e6a-146">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c1e6a-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c1e6a-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c1e6a-147">ReadItem</span></span>|
|[<span data-ttu-id="c1e6a-148">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c1e6a-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c1e6a-149">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c1e6a-149">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c1e6a-150">Exemple</span><span class="sxs-lookup"><span data-stu-id="c1e6a-150">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="c1e6a-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="c1e6a-151">timeZone :String</span></span>

<span data-ttu-id="c1e6a-152">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="c1e6a-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="c1e6a-153">Type :</span><span class="sxs-lookup"><span data-stu-id="c1e6a-153">Type:</span></span>

*   <span data-ttu-id="c1e6a-154">String</span><span class="sxs-lookup"><span data-stu-id="c1e6a-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c1e6a-155">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c1e6a-155">Requirements</span></span>

|<span data-ttu-id="c1e6a-156">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c1e6a-156">Requirement</span></span>| <span data-ttu-id="c1e6a-157">Valeur</span><span class="sxs-lookup"><span data-stu-id="c1e6a-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="c1e6a-158">Version minimale de l’ensemble de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c1e6a-158">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c1e6a-159">1.0</span><span class="sxs-lookup"><span data-stu-id="c1e6a-159">1.0</span></span>|
|[<span data-ttu-id="c1e6a-160">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c1e6a-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c1e6a-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c1e6a-161">ReadItem</span></span>|
|[<span data-ttu-id="c1e6a-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c1e6a-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c1e6a-163">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c1e6a-163">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c1e6a-164">Exemple</span><span class="sxs-lookup"><span data-stu-id="c1e6a-164">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```