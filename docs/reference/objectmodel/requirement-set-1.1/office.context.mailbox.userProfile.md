
# <a name="userprofile"></a><span data-ttu-id="de5b7-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="de5b7-101">userProfile</span></span>

### <span data-ttu-id="de5b7-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="de5b7-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="de5b7-104">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="de5b7-104">Requirements</span></span>

|<span data-ttu-id="de5b7-105">Condition requise</span><span class="sxs-lookup"><span data-stu-id="de5b7-105">Requirement</span></span>| <span data-ttu-id="de5b7-106">Valeur</span><span class="sxs-lookup"><span data-stu-id="de5b7-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="de5b7-107">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="de5b7-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="de5b7-108">1.0</span><span class="sxs-lookup"><span data-stu-id="de5b7-108">1.0</span></span>|
|[<span data-ttu-id="de5b7-109">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="de5b7-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="de5b7-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="de5b7-110">ReadItem</span></span>|
|[<span data-ttu-id="de5b7-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="de5b7-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="de5b7-112">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="de5b7-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="de5b7-113">Membres</span><span class="sxs-lookup"><span data-stu-id="de5b7-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="de5b7-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="de5b7-114">displayName :String</span></span>

<span data-ttu-id="de5b7-115">Obtient le nom d’affichage de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="de5b7-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="de5b7-116">Type :</span><span class="sxs-lookup"><span data-stu-id="de5b7-116">Type:</span></span>

*   <span data-ttu-id="de5b7-117">Chaîne</span><span class="sxs-lookup"><span data-stu-id="de5b7-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="de5b7-118">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="de5b7-118">Requirements</span></span>

|<span data-ttu-id="de5b7-119">Condition requise</span><span class="sxs-lookup"><span data-stu-id="de5b7-119">Requirement</span></span>| <span data-ttu-id="de5b7-120">Valeur</span><span class="sxs-lookup"><span data-stu-id="de5b7-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="de5b7-121">Version minimale de l’ensemble des conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="de5b7-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="de5b7-122">1.0</span><span class="sxs-lookup"><span data-stu-id="de5b7-122">1.0</span></span>|
|[<span data-ttu-id="de5b7-123">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="de5b7-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="de5b7-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="de5b7-124">ReadItem</span></span>|
|[<span data-ttu-id="de5b7-125">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="de5b7-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="de5b7-126">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="de5b7-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="de5b7-127">Exemple</span><span class="sxs-lookup"><span data-stu-id="de5b7-127">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="de5b7-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="de5b7-128">emailAddress :String</span></span>

<span data-ttu-id="de5b7-129">Obtient l’adresse e-mail SMTP de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="de5b7-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="de5b7-130">Type :</span><span class="sxs-lookup"><span data-stu-id="de5b7-130">Type:</span></span>

*   <span data-ttu-id="de5b7-131">Chaîne</span><span class="sxs-lookup"><span data-stu-id="de5b7-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="de5b7-132">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="de5b7-132">Requirements</span></span>

|<span data-ttu-id="de5b7-133">Condition requise</span><span class="sxs-lookup"><span data-stu-id="de5b7-133">Requirement</span></span>| <span data-ttu-id="de5b7-134">Valeur</span><span class="sxs-lookup"><span data-stu-id="de5b7-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="de5b7-135">Version minimale des exigences de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="de5b7-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="de5b7-136">1.0</span><span class="sxs-lookup"><span data-stu-id="de5b7-136">1.0</span></span>|
|[<span data-ttu-id="de5b7-137">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="de5b7-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="de5b7-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="de5b7-138">ReadItem</span></span>|
|[<span data-ttu-id="de5b7-139">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="de5b7-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="de5b7-140">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="de5b7-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="de5b7-141">Exemple</span><span class="sxs-lookup"><span data-stu-id="de5b7-141">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="de5b7-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="de5b7-142">timeZone :String</span></span>

<span data-ttu-id="de5b7-143">Obtient le fuseau horaire par défaut de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="de5b7-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="de5b7-144">Type :</span><span class="sxs-lookup"><span data-stu-id="de5b7-144">Type:</span></span>

*   <span data-ttu-id="de5b7-145">Chaîne</span><span class="sxs-lookup"><span data-stu-id="de5b7-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="de5b7-146">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="de5b7-146">Requirements</span></span>

|<span data-ttu-id="de5b7-147">Condition requise</span><span class="sxs-lookup"><span data-stu-id="de5b7-147">Requirement</span></span>| <span data-ttu-id="de5b7-148">Valeur</span><span class="sxs-lookup"><span data-stu-id="de5b7-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="de5b7-149">Version minimale de l’ensemble des conditions requises de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="de5b7-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="de5b7-150">1.0</span><span class="sxs-lookup"><span data-stu-id="de5b7-150">1.0</span></span>|
|[<span data-ttu-id="de5b7-151">Niveau minimal d’autorisation</span><span class="sxs-lookup"><span data-stu-id="de5b7-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="de5b7-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="de5b7-152">ReadItem</span></span>|
|[<span data-ttu-id="de5b7-153">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="de5b7-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="de5b7-154">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="de5b7-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="de5b7-155">Exemple</span><span class="sxs-lookup"><span data-stu-id="de5b7-155">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```