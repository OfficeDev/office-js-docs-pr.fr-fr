
# <a name="userprofile"></a><span data-ttu-id="457bc-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="457bc-101">userProfile</span></span>

### <span data-ttu-id="457bc-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="457bc-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="457bc-104">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="457bc-104">Requirements</span></span>

|<span data-ttu-id="457bc-105">Condition requise</span><span class="sxs-lookup"><span data-stu-id="457bc-105">Requirement</span></span>| <span data-ttu-id="457bc-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="457bc-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="457bc-107">Version minimale de l’ensemble de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="457bc-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="457bc-108">1.0</span><span class="sxs-lookup"><span data-stu-id="457bc-108">1.0</span></span>|
|[<span data-ttu-id="457bc-109">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="457bc-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="457bc-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="457bc-110">ReadItem</span></span>|
|[<span data-ttu-id="457bc-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="457bc-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="457bc-112">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="457bc-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="457bc-113">Membres</span><span class="sxs-lookup"><span data-stu-id="457bc-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="457bc-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="457bc-114">displayName :String</span></span>

<span data-ttu-id="457bc-115">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="457bc-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="457bc-116">Type :</span><span class="sxs-lookup"><span data-stu-id="457bc-116">Type:</span></span>

*   <span data-ttu-id="457bc-117">String</span><span class="sxs-lookup"><span data-stu-id="457bc-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="457bc-118">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="457bc-118">Requirements</span></span>

|<span data-ttu-id="457bc-119">Condition requise</span><span class="sxs-lookup"><span data-stu-id="457bc-119">Requirement</span></span>| <span data-ttu-id="457bc-120">Valeur</span><span class="sxs-lookup"><span data-stu-id="457bc-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="457bc-121">Version minimale de l’ensemble de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="457bc-121">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="457bc-122">1.0</span><span class="sxs-lookup"><span data-stu-id="457bc-122">1.0</span></span>|
|[<span data-ttu-id="457bc-123">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="457bc-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="457bc-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="457bc-124">ReadItem</span></span>|
|[<span data-ttu-id="457bc-125">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="457bc-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="457bc-126">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="457bc-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="457bc-127">Exemple</span><span class="sxs-lookup"><span data-stu-id="457bc-127">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="457bc-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="457bc-128">emailAddress :String</span></span>

<span data-ttu-id="457bc-129">Obtient l’adresse e-mail SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="457bc-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="457bc-130">Type :</span><span class="sxs-lookup"><span data-stu-id="457bc-130">Type:</span></span>

*   <span data-ttu-id="457bc-131">String</span><span class="sxs-lookup"><span data-stu-id="457bc-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="457bc-132">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="457bc-132">Requirements</span></span>

|<span data-ttu-id="457bc-133">Condition requise</span><span class="sxs-lookup"><span data-stu-id="457bc-133">Requirement</span></span>| <span data-ttu-id="457bc-134">Valeur</span><span class="sxs-lookup"><span data-stu-id="457bc-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="457bc-135">Version minimale de l’ensemble de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="457bc-135">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="457bc-136">1.0</span><span class="sxs-lookup"><span data-stu-id="457bc-136">1.0</span></span>|
|[<span data-ttu-id="457bc-137">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="457bc-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="457bc-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="457bc-138">ReadItem</span></span>|
|[<span data-ttu-id="457bc-139">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="457bc-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="457bc-140">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="457bc-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="457bc-141">Exemple</span><span class="sxs-lookup"><span data-stu-id="457bc-141">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="457bc-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="457bc-142">timeZone :String</span></span>

<span data-ttu-id="457bc-143">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="457bc-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="457bc-144">Type :</span><span class="sxs-lookup"><span data-stu-id="457bc-144">Type:</span></span>

*   <span data-ttu-id="457bc-145">String</span><span class="sxs-lookup"><span data-stu-id="457bc-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="457bc-146">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="457bc-146">Requirements</span></span>

|<span data-ttu-id="457bc-147">Condition requise</span><span class="sxs-lookup"><span data-stu-id="457bc-147">Requirement</span></span>| <span data-ttu-id="457bc-148">Valeur</span><span class="sxs-lookup"><span data-stu-id="457bc-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="457bc-149">Version minimale de l’ensemble de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="457bc-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="457bc-150">1.0</span><span class="sxs-lookup"><span data-stu-id="457bc-150">1.0</span></span>|
|[<span data-ttu-id="457bc-151">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="457bc-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="457bc-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="457bc-152">ReadItem</span></span>|
|[<span data-ttu-id="457bc-153">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="457bc-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="457bc-154">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="457bc-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="457bc-155">Exemple</span><span class="sxs-lookup"><span data-stu-id="457bc-155">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```