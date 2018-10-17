
# <a name="userprofile"></a><span data-ttu-id="6caae-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="6caae-101">userProfile</span></span>

### <span data-ttu-id="6caae-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="6caae-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="6caae-104">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6caae-104">Requirements</span></span>

|<span data-ttu-id="6caae-105">Condition requise</span><span class="sxs-lookup"><span data-stu-id="6caae-105">Requirement</span></span>| <span data-ttu-id="6caae-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="6caae-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="6caae-107">Version minimale requise de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6caae-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6caae-108">1.0</span><span class="sxs-lookup"><span data-stu-id="6caae-108">1.0</span></span>|
|[<span data-ttu-id="6caae-109">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="6caae-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6caae-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6caae-110">ReadItem</span></span>|
|[<span data-ttu-id="6caae-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6caae-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6caae-112">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="6caae-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="6caae-113">Membres</span><span class="sxs-lookup"><span data-stu-id="6caae-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="6caae-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="6caae-114">displayName :String</span></span>

<span data-ttu-id="6caae-115">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="6caae-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="6caae-116">Type :</span><span class="sxs-lookup"><span data-stu-id="6caae-116">Type:</span></span>

*   <span data-ttu-id="6caae-117">Chaîne</span><span class="sxs-lookup"><span data-stu-id="6caae-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6caae-118">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6caae-118">Requirements</span></span>

|<span data-ttu-id="6caae-119">Condition requise</span><span class="sxs-lookup"><span data-stu-id="6caae-119">Requirement</span></span>| <span data-ttu-id="6caae-120">Valeur</span><span class="sxs-lookup"><span data-stu-id="6caae-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="6caae-121">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6caae-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6caae-122">1.0</span><span class="sxs-lookup"><span data-stu-id="6caae-122">1.0</span></span>|
|[<span data-ttu-id="6caae-123">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="6caae-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6caae-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6caae-124">ReadItem</span></span>|
|[<span data-ttu-id="6caae-125">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6caae-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6caae-126">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="6caae-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6caae-127">Exemple</span><span class="sxs-lookup"><span data-stu-id="6caae-127">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="6caae-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="6caae-128">emailAddress :String</span></span>

<span data-ttu-id="6caae-129">Obtient l’adresse e-mail SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="6caae-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="6caae-130">Type :</span><span class="sxs-lookup"><span data-stu-id="6caae-130">Type:</span></span>

*   <span data-ttu-id="6caae-131">Chaîne</span><span class="sxs-lookup"><span data-stu-id="6caae-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6caae-132">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6caae-132">Requirements</span></span>

|<span data-ttu-id="6caae-133">Condition requise</span><span class="sxs-lookup"><span data-stu-id="6caae-133">Requirement</span></span>| <span data-ttu-id="6caae-134">Valeur</span><span class="sxs-lookup"><span data-stu-id="6caae-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="6caae-135">Version minimale requise de la boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6caae-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6caae-136">1.0</span><span class="sxs-lookup"><span data-stu-id="6caae-136">1.0</span></span>|
|[<span data-ttu-id="6caae-137">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="6caae-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6caae-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6caae-138">ReadItem</span></span>|
|[<span data-ttu-id="6caae-139">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6caae-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6caae-140">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="6caae-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6caae-141">Exemple</span><span class="sxs-lookup"><span data-stu-id="6caae-141">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="6caae-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="6caae-142">timeZone :String</span></span>

<span data-ttu-id="6caae-143">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="6caae-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="6caae-144">Type :</span><span class="sxs-lookup"><span data-stu-id="6caae-144">Type:</span></span>

*   <span data-ttu-id="6caae-145">Chaîne</span><span class="sxs-lookup"><span data-stu-id="6caae-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6caae-146">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6caae-146">Requirements</span></span>

|<span data-ttu-id="6caae-147">Condition requise</span><span class="sxs-lookup"><span data-stu-id="6caae-147">Requirement</span></span>| <span data-ttu-id="6caae-148">Valeur</span><span class="sxs-lookup"><span data-stu-id="6caae-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="6caae-149">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6caae-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6caae-150">1.0</span><span class="sxs-lookup"><span data-stu-id="6caae-150">1.0</span></span>|
|[<span data-ttu-id="6caae-151">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="6caae-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6caae-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6caae-152">ReadItem</span></span>|
|[<span data-ttu-id="6caae-153">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6caae-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6caae-154">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="6caae-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6caae-155">Exemple</span><span class="sxs-lookup"><span data-stu-id="6caae-155">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```