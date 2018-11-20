
# <a name="userprofile"></a><span data-ttu-id="fbab3-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="fbab3-101">userProfile</span></span>

### <span data-ttu-id="fbab3-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="fbab3-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="fbab3-104">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fbab3-104">Requirements</span></span>

|<span data-ttu-id="fbab3-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fbab3-105">Requirement</span></span>| <span data-ttu-id="fbab3-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="fbab3-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="fbab3-107">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fbab3-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fbab3-108">1.0</span><span class="sxs-lookup"><span data-stu-id="fbab3-108">1.0</span></span>|
|[<span data-ttu-id="fbab3-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fbab3-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fbab3-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fbab3-110">ReadItem</span></span>|
|[<span data-ttu-id="fbab3-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fbab3-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fbab3-112">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="fbab3-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="fbab3-113">Membres</span><span class="sxs-lookup"><span data-stu-id="fbab3-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="fbab3-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="fbab3-114">displayName :String</span></span>

<span data-ttu-id="fbab3-115">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="fbab3-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="fbab3-116">Type :</span><span class="sxs-lookup"><span data-stu-id="fbab3-116">Type:</span></span>

*   <span data-ttu-id="fbab3-117">Chaîne</span><span class="sxs-lookup"><span data-stu-id="fbab3-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fbab3-118">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fbab3-118">Requirements</span></span>

|<span data-ttu-id="fbab3-119">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fbab3-119">Requirement</span></span>| <span data-ttu-id="fbab3-120">Valeur</span><span class="sxs-lookup"><span data-stu-id="fbab3-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="fbab3-121">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fbab3-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fbab3-122">1.0</span><span class="sxs-lookup"><span data-stu-id="fbab3-122">1.0</span></span>|
|[<span data-ttu-id="fbab3-123">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fbab3-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fbab3-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fbab3-124">ReadItem</span></span>|
|[<span data-ttu-id="fbab3-125">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fbab3-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fbab3-126">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="fbab3-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fbab3-127">Exemple</span><span class="sxs-lookup"><span data-stu-id="fbab3-127">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="fbab3-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="fbab3-128">emailAddress :String</span></span>

<span data-ttu-id="fbab3-129">Obtient l’adresse de messagerie SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="fbab3-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="fbab3-130">Type :</span><span class="sxs-lookup"><span data-stu-id="fbab3-130">Type:</span></span>

*   <span data-ttu-id="fbab3-131">Chaîne</span><span class="sxs-lookup"><span data-stu-id="fbab3-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fbab3-132">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fbab3-132">Requirements</span></span>

|<span data-ttu-id="fbab3-133">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fbab3-133">Requirement</span></span>| <span data-ttu-id="fbab3-134">Valeur</span><span class="sxs-lookup"><span data-stu-id="fbab3-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="fbab3-135">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fbab3-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fbab3-136">1.0</span><span class="sxs-lookup"><span data-stu-id="fbab3-136">1.0</span></span>|
|[<span data-ttu-id="fbab3-137">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fbab3-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fbab3-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fbab3-138">ReadItem</span></span>|
|[<span data-ttu-id="fbab3-139">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fbab3-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fbab3-140">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="fbab3-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fbab3-141">Exemple</span><span class="sxs-lookup"><span data-stu-id="fbab3-141">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="fbab3-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="fbab3-142">timeZone :String</span></span>

<span data-ttu-id="fbab3-143">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="fbab3-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="fbab3-144">Type :</span><span class="sxs-lookup"><span data-stu-id="fbab3-144">Type:</span></span>

*   <span data-ttu-id="fbab3-145">Chaîne</span><span class="sxs-lookup"><span data-stu-id="fbab3-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="fbab3-146">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="fbab3-146">Requirements</span></span>

|<span data-ttu-id="fbab3-147">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="fbab3-147">Requirement</span></span>| <span data-ttu-id="fbab3-148">Valeur</span><span class="sxs-lookup"><span data-stu-id="fbab3-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="fbab3-149">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="fbab3-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="fbab3-150">1.0</span><span class="sxs-lookup"><span data-stu-id="fbab3-150">1.0</span></span>|
|[<span data-ttu-id="fbab3-151">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="fbab3-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="fbab3-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="fbab3-152">ReadItem</span></span>|
|[<span data-ttu-id="fbab3-153">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="fbab3-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="fbab3-154">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="fbab3-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="fbab3-155">Exemple</span><span class="sxs-lookup"><span data-stu-id="fbab3-155">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```