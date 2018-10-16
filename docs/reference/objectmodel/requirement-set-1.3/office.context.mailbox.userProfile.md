
# <a name="userprofile"></a><span data-ttu-id="c82c4-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="c82c4-101">userProfile</span></span>

### <span data-ttu-id="c82c4-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="c82c4-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="c82c4-104">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c82c4-104">Requirements</span></span>

|<span data-ttu-id="c82c4-105">Condition requise</span><span class="sxs-lookup"><span data-stu-id="c82c4-105">Requirement</span></span>| <span data-ttu-id="c82c4-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="c82c4-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="c82c4-107">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c82c4-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c82c4-108">1.0</span><span class="sxs-lookup"><span data-stu-id="c82c4-108">1.0</span></span>|
|[<span data-ttu-id="c82c4-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c82c4-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c82c4-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c82c4-110">ReadItem</span></span>|
|[<span data-ttu-id="c82c4-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c82c4-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c82c4-112">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c82c4-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="c82c4-113">Membres</span><span class="sxs-lookup"><span data-stu-id="c82c4-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="c82c4-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="c82c4-114">displayName :String</span></span>

<span data-ttu-id="c82c4-115">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="c82c4-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="c82c4-116">Type :</span><span class="sxs-lookup"><span data-stu-id="c82c4-116">Type:</span></span>

*   <span data-ttu-id="c82c4-117">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c82c4-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c82c4-118">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c82c4-118">Requirements</span></span>

|<span data-ttu-id="c82c4-119">Condition requise</span><span class="sxs-lookup"><span data-stu-id="c82c4-119">Requirement</span></span>| <span data-ttu-id="c82c4-120">Valeur</span><span class="sxs-lookup"><span data-stu-id="c82c4-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="c82c4-121">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c82c4-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c82c4-122">1.0</span><span class="sxs-lookup"><span data-stu-id="c82c4-122">1.0</span></span>|
|[<span data-ttu-id="c82c4-123">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c82c4-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c82c4-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c82c4-124">ReadItem</span></span>|
|[<span data-ttu-id="c82c4-125">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c82c4-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c82c4-126">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c82c4-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c82c4-127">Exemple</span><span class="sxs-lookup"><span data-stu-id="c82c4-127">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="c82c4-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="c82c4-128">emailAddress :String</span></span>

<span data-ttu-id="c82c4-129">Obtient l’adresse e-mail SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="c82c4-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="c82c4-130">Type :</span><span class="sxs-lookup"><span data-stu-id="c82c4-130">Type:</span></span>

*   <span data-ttu-id="c82c4-131">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c82c4-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c82c4-132">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c82c4-132">Requirements</span></span>

|<span data-ttu-id="c82c4-133">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c82c4-133">Requirement</span></span>| <span data-ttu-id="c82c4-134">Valeur</span><span class="sxs-lookup"><span data-stu-id="c82c4-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="c82c4-135">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c82c4-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c82c4-136">1.0</span><span class="sxs-lookup"><span data-stu-id="c82c4-136">1.0</span></span>|
|[<span data-ttu-id="c82c4-137">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c82c4-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c82c4-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c82c4-138">ReadItem</span></span>|
|[<span data-ttu-id="c82c4-139">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c82c4-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c82c4-140">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c82c4-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c82c4-141">Exemple</span><span class="sxs-lookup"><span data-stu-id="c82c4-141">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="c82c4-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="c82c4-142">timeZone :String</span></span>

<span data-ttu-id="c82c4-143">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="c82c4-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="c82c4-144">Type :</span><span class="sxs-lookup"><span data-stu-id="c82c4-144">Type:</span></span>

*   <span data-ttu-id="c82c4-145">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c82c4-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c82c4-146">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c82c4-146">Requirements</span></span>

|<span data-ttu-id="c82c4-147">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c82c4-147">Requirement</span></span>| <span data-ttu-id="c82c4-148">Valeur</span><span class="sxs-lookup"><span data-stu-id="c82c4-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="c82c4-149">Version minimale de l’ensemble des conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c82c4-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c82c4-150">1.0</span><span class="sxs-lookup"><span data-stu-id="c82c4-150">1.0</span></span>|
|[<span data-ttu-id="c82c4-151">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c82c4-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c82c4-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c82c4-152">ReadItem</span></span>|
|[<span data-ttu-id="c82c4-153">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c82c4-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c82c4-154">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c82c4-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c82c4-155">Exemple</span><span class="sxs-lookup"><span data-stu-id="c82c4-155">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```