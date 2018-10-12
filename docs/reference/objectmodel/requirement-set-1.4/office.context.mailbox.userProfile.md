
# <a name="userprofile"></a><span data-ttu-id="0fd9c-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="0fd9c-101">userProfile</span></span>

### <span data-ttu-id="0fd9c-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="0fd9c-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="0fd9c-104">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0fd9c-104">Requirements</span></span>

|<span data-ttu-id="0fd9c-105">Condition requise</span><span class="sxs-lookup"><span data-stu-id="0fd9c-105">Requirement</span></span>| <span data-ttu-id="0fd9c-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="0fd9c-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="0fd9c-107">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0fd9c-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0fd9c-108">1.0</span><span class="sxs-lookup"><span data-stu-id="0fd9c-108">1.0</span></span>|
|[<span data-ttu-id="0fd9c-109">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="0fd9c-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0fd9c-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0fd9c-110">ReadItem</span></span>|
|[<span data-ttu-id="0fd9c-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0fd9c-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0fd9c-112">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="0fd9c-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="0fd9c-113">Membres</span><span class="sxs-lookup"><span data-stu-id="0fd9c-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="0fd9c-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="0fd9c-114">displayName :String</span></span>

<span data-ttu-id="0fd9c-115">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="0fd9c-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="0fd9c-116">Type :</span><span class="sxs-lookup"><span data-stu-id="0fd9c-116">Type:</span></span>

*   <span data-ttu-id="0fd9c-117">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0fd9c-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0fd9c-118">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0fd9c-118">Requirements</span></span>

|<span data-ttu-id="0fd9c-119">Condition requise</span><span class="sxs-lookup"><span data-stu-id="0fd9c-119">Requirement</span></span>| <span data-ttu-id="0fd9c-120">Valeur</span><span class="sxs-lookup"><span data-stu-id="0fd9c-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="0fd9c-121">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0fd9c-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0fd9c-122">1.0</span><span class="sxs-lookup"><span data-stu-id="0fd9c-122">1.0</span></span>|
|[<span data-ttu-id="0fd9c-123">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="0fd9c-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0fd9c-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0fd9c-124">ReadItem</span></span>|
|[<span data-ttu-id="0fd9c-125">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0fd9c-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0fd9c-126">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="0fd9c-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="0fd9c-127">Exemple</span><span class="sxs-lookup"><span data-stu-id="0fd9c-127">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="0fd9c-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="0fd9c-128">emailAddress :String</span></span>

<span data-ttu-id="0fd9c-129">Obtient l’adresse e-mail SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="0fd9c-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="0fd9c-130">Type :</span><span class="sxs-lookup"><span data-stu-id="0fd9c-130">Type:</span></span>

*   <span data-ttu-id="0fd9c-131">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0fd9c-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0fd9c-132">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0fd9c-132">Requirements</span></span>

|<span data-ttu-id="0fd9c-133">Condition requise</span><span class="sxs-lookup"><span data-stu-id="0fd9c-133">Requirement</span></span>| <span data-ttu-id="0fd9c-134">Valeur</span><span class="sxs-lookup"><span data-stu-id="0fd9c-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="0fd9c-135">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0fd9c-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0fd9c-136">1.0</span><span class="sxs-lookup"><span data-stu-id="0fd9c-136">1.0</span></span>|
|[<span data-ttu-id="0fd9c-137">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="0fd9c-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0fd9c-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0fd9c-138">ReadItem</span></span>|
|[<span data-ttu-id="0fd9c-139">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0fd9c-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0fd9c-140">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="0fd9c-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="0fd9c-141">Exemple</span><span class="sxs-lookup"><span data-stu-id="0fd9c-141">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="0fd9c-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="0fd9c-142">timeZone :String</span></span>

<span data-ttu-id="0fd9c-143">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="0fd9c-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="0fd9c-144">Type :</span><span class="sxs-lookup"><span data-stu-id="0fd9c-144">Type:</span></span>

*   <span data-ttu-id="0fd9c-145">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0fd9c-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0fd9c-146">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0fd9c-146">Requirements</span></span>

|<span data-ttu-id="0fd9c-147">Condition requise</span><span class="sxs-lookup"><span data-stu-id="0fd9c-147">Requirement</span></span>| <span data-ttu-id="0fd9c-148">Valeur</span><span class="sxs-lookup"><span data-stu-id="0fd9c-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="0fd9c-149">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0fd9c-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0fd9c-150">1.0</span><span class="sxs-lookup"><span data-stu-id="0fd9c-150">1.0</span></span>|
|[<span data-ttu-id="0fd9c-151">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="0fd9c-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0fd9c-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0fd9c-152">ReadItem</span></span>|
|[<span data-ttu-id="0fd9c-153">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0fd9c-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0fd9c-154">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="0fd9c-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="0fd9c-155">Exemple</span><span class="sxs-lookup"><span data-stu-id="0fd9c-155">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```