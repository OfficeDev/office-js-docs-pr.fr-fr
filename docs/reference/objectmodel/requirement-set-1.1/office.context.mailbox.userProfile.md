
# <a name="userprofile"></a><span data-ttu-id="a04fb-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="a04fb-101">userProfile</span></span>

### <span data-ttu-id="a04fb-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="a04fb-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="a04fb-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a04fb-104">Requirements</span></span>

|<span data-ttu-id="a04fb-105">Condition requise</span><span class="sxs-lookup"><span data-stu-id="a04fb-105">Requirement</span></span>| <span data-ttu-id="a04fb-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="a04fb-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="a04fb-107">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a04fb-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a04fb-108">1.0</span><span class="sxs-lookup"><span data-stu-id="a04fb-108">1.0</span></span>|
|[<span data-ttu-id="a04fb-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a04fb-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a04fb-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a04fb-110">ReadItem</span></span>|
|[<span data-ttu-id="a04fb-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a04fb-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a04fb-112">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="a04fb-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="a04fb-113">Membres</span><span class="sxs-lookup"><span data-stu-id="a04fb-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="a04fb-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="a04fb-114">displayName :String</span></span>

<span data-ttu-id="a04fb-115">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="a04fb-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="a04fb-116">Type :</span><span class="sxs-lookup"><span data-stu-id="a04fb-116">Type:</span></span>

*   <span data-ttu-id="a04fb-117">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a04fb-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a04fb-118">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a04fb-118">Requirements</span></span>

|<span data-ttu-id="a04fb-119">Condition requise</span><span class="sxs-lookup"><span data-stu-id="a04fb-119">Requirement</span></span>| <span data-ttu-id="a04fb-120">Valeur</span><span class="sxs-lookup"><span data-stu-id="a04fb-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="a04fb-121">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a04fb-121">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a04fb-122">1.0</span><span class="sxs-lookup"><span data-stu-id="a04fb-122">1.0</span></span>|
|[<span data-ttu-id="a04fb-123">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a04fb-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a04fb-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a04fb-124">ReadItem</span></span>|
|[<span data-ttu-id="a04fb-125">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a04fb-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a04fb-126">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="a04fb-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a04fb-127">Exemple</span><span class="sxs-lookup"><span data-stu-id="a04fb-127">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="a04fb-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="a04fb-128">emailAddress :String</span></span>

<span data-ttu-id="a04fb-129">Obtient l’adresse e-mail SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="a04fb-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="a04fb-130">Type :</span><span class="sxs-lookup"><span data-stu-id="a04fb-130">Type:</span></span>

*   <span data-ttu-id="a04fb-131">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a04fb-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a04fb-132">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a04fb-132">Requirements</span></span>

|<span data-ttu-id="a04fb-133">Condition requise</span><span class="sxs-lookup"><span data-stu-id="a04fb-133">Requirement</span></span>| <span data-ttu-id="a04fb-134">Valeur</span><span class="sxs-lookup"><span data-stu-id="a04fb-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="a04fb-135">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a04fb-135">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a04fb-136">1.0</span><span class="sxs-lookup"><span data-stu-id="a04fb-136">1.0</span></span>|
|[<span data-ttu-id="a04fb-137">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a04fb-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a04fb-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a04fb-138">ReadItem</span></span>|
|[<span data-ttu-id="a04fb-139">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a04fb-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a04fb-140">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="a04fb-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a04fb-141">Exemple</span><span class="sxs-lookup"><span data-stu-id="a04fb-141">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="a04fb-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="a04fb-142">timeZone :String</span></span>

<span data-ttu-id="a04fb-143">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="a04fb-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="a04fb-144">Type :</span><span class="sxs-lookup"><span data-stu-id="a04fb-144">Type:</span></span>

*   <span data-ttu-id="a04fb-145">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a04fb-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a04fb-146">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a04fb-146">Requirements</span></span>

|<span data-ttu-id="a04fb-147">Condition requise</span><span class="sxs-lookup"><span data-stu-id="a04fb-147">Requirement</span></span>| <span data-ttu-id="a04fb-148">Valeur</span><span class="sxs-lookup"><span data-stu-id="a04fb-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="a04fb-149">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a04fb-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a04fb-150">1.0</span><span class="sxs-lookup"><span data-stu-id="a04fb-150">1.0</span></span>|
|[<span data-ttu-id="a04fb-151">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a04fb-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a04fb-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a04fb-152">ReadItem</span></span>|
|[<span data-ttu-id="a04fb-153">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a04fb-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a04fb-154">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="a04fb-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a04fb-155">Exemple</span><span class="sxs-lookup"><span data-stu-id="a04fb-155">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```