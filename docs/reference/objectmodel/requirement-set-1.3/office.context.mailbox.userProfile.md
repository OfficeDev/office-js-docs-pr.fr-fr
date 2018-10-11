
# <a name="userprofile"></a><span data-ttu-id="2c278-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="2c278-101">userProfile</span></span>

### <span data-ttu-id="2c278-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="2c278-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="2c278-104">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2c278-104">Requirements</span></span>

|<span data-ttu-id="2c278-105">Condition requise</span><span class="sxs-lookup"><span data-stu-id="2c278-105">Requirement</span></span>| <span data-ttu-id="2c278-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="2c278-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="2c278-107">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2c278-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2c278-108">1.0</span><span class="sxs-lookup"><span data-stu-id="2c278-108">1.0</span></span>|
|[<span data-ttu-id="2c278-109">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2c278-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2c278-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2c278-110">ReadItem</span></span>|
|[<span data-ttu-id="2c278-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2c278-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2c278-112">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="2c278-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="2c278-113">Membres</span><span class="sxs-lookup"><span data-stu-id="2c278-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="2c278-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="2c278-114">displayName :String</span></span>

<span data-ttu-id="2c278-115">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="2c278-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="2c278-116">Type :</span><span class="sxs-lookup"><span data-stu-id="2c278-116">Type:</span></span>

*   <span data-ttu-id="2c278-117">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2c278-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2c278-118">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2c278-118">Requirements</span></span>

|<span data-ttu-id="2c278-119">Condition requise</span><span class="sxs-lookup"><span data-stu-id="2c278-119">Requirement</span></span>| <span data-ttu-id="2c278-120">Valeur</span><span class="sxs-lookup"><span data-stu-id="2c278-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="2c278-121">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2c278-121">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2c278-122">1.0</span><span class="sxs-lookup"><span data-stu-id="2c278-122">1.0</span></span>|
|[<span data-ttu-id="2c278-123">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2c278-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2c278-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2c278-124">ReadItem</span></span>|
|[<span data-ttu-id="2c278-125">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2c278-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2c278-126">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="2c278-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2c278-127">Exemple</span><span class="sxs-lookup"><span data-stu-id="2c278-127">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="2c278-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="2c278-128">emailAddress :String</span></span>

<span data-ttu-id="2c278-129">Obtient l’adresse e-mail SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="2c278-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="2c278-130">Type :</span><span class="sxs-lookup"><span data-stu-id="2c278-130">Type:</span></span>

*   <span data-ttu-id="2c278-131">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2c278-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2c278-132">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2c278-132">Requirements</span></span>

|<span data-ttu-id="2c278-133">Condition requise</span><span class="sxs-lookup"><span data-stu-id="2c278-133">Requirement</span></span>| <span data-ttu-id="2c278-134">Valeur</span><span class="sxs-lookup"><span data-stu-id="2c278-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="2c278-135">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2c278-135">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2c278-136">1.0</span><span class="sxs-lookup"><span data-stu-id="2c278-136">1.0</span></span>|
|[<span data-ttu-id="2c278-137">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2c278-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2c278-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2c278-138">ReadItem</span></span>|
|[<span data-ttu-id="2c278-139">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2c278-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2c278-140">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="2c278-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2c278-141">Exemple</span><span class="sxs-lookup"><span data-stu-id="2c278-141">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="2c278-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="2c278-142">timeZone :String</span></span>

<span data-ttu-id="2c278-143">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="2c278-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="2c278-144">Type :</span><span class="sxs-lookup"><span data-stu-id="2c278-144">Type:</span></span>

*   <span data-ttu-id="2c278-145">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2c278-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2c278-146">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2c278-146">Requirements</span></span>

|<span data-ttu-id="2c278-147">Condition requise</span><span class="sxs-lookup"><span data-stu-id="2c278-147">Requirement</span></span>| <span data-ttu-id="2c278-148">Valeur</span><span class="sxs-lookup"><span data-stu-id="2c278-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="2c278-149">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2c278-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2c278-150">1.0</span><span class="sxs-lookup"><span data-stu-id="2c278-150">1.0</span></span>|
|[<span data-ttu-id="2c278-151">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2c278-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2c278-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2c278-152">ReadItem</span></span>|
|[<span data-ttu-id="2c278-153">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2c278-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2c278-154">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="2c278-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2c278-155">Exemple</span><span class="sxs-lookup"><span data-stu-id="2c278-155">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```