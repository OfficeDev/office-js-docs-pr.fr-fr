
# <a name="diagnostics"></a><span data-ttu-id="6b2a8-101">diagnostics</span><span class="sxs-lookup"><span data-stu-id="6b2a8-101">diagnostics</span></span>

### <span data-ttu-id="6b2a8-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics</span><span class="sxs-lookup"><span data-stu-id="6b2a8-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). diagnostics</span></span>

<span data-ttu-id="6b2a8-104">Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="6b2a8-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="6b2a8-105">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6b2a8-105">Requirements</span></span>

|<span data-ttu-id="6b2a8-106">Condition requise</span><span class="sxs-lookup"><span data-stu-id="6b2a8-106">Requirement</span></span>| <span data-ttu-id="6b2a8-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="6b2a8-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="6b2a8-108">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6b2a8-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6b2a8-109">1.0</span><span class="sxs-lookup"><span data-stu-id="6b2a8-109">1.0</span></span>|
|[<span data-ttu-id="6b2a8-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6b2a8-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6b2a8-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6b2a8-111">ReadItem</span></span>|
|[<span data-ttu-id="6b2a8-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6b2a8-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6b2a8-113">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="6b2a8-113">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="6b2a8-114">Membres</span><span class="sxs-lookup"><span data-stu-id="6b2a8-114">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="6b2a8-115">hostName :String</span><span class="sxs-lookup"><span data-stu-id="6b2a8-115">hostName :String</span></span>

<span data-ttu-id="6b2a8-116">Obtient une chaîne qui représente le nom de l’application hôte.</span><span class="sxs-lookup"><span data-stu-id="6b2a8-116">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="6b2a8-117">Chaîne qui peut avoir l’une des valeurs suivantes : `Outlook`, `OutlookIOS` ou `OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="6b2a8-117">A string that can be one of the following values: `Outlook`, `OutlookIOS`, , or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="6b2a8-118">Type :</span><span class="sxs-lookup"><span data-stu-id="6b2a8-118">Type:</span></span>

*   <span data-ttu-id="6b2a8-119">Chaîne</span><span class="sxs-lookup"><span data-stu-id="6b2a8-119">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6b2a8-120">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6b2a8-120">Requirements</span></span>

|<span data-ttu-id="6b2a8-121">Condition requise</span><span class="sxs-lookup"><span data-stu-id="6b2a8-121">Requirement</span></span>| <span data-ttu-id="6b2a8-122">Valeur</span><span class="sxs-lookup"><span data-stu-id="6b2a8-122">Value</span></span>|
|---|---|
|[<span data-ttu-id="6b2a8-123">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6b2a8-123">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6b2a8-124">1.0</span><span class="sxs-lookup"><span data-stu-id="6b2a8-124">1.0</span></span>|
|[<span data-ttu-id="6b2a8-125">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6b2a8-125">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6b2a8-126">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6b2a8-126">ReadItem</span></span>|
|[<span data-ttu-id="6b2a8-127">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6b2a8-127">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6b2a8-128">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="6b2a8-128">Compose or read</span></span>|

####  <a name="hostversion-string"></a><span data-ttu-id="6b2a8-129">hostVersion :String</span><span class="sxs-lookup"><span data-stu-id="6b2a8-129">hostVersion :String</span></span>

<span data-ttu-id="6b2a8-130">Obtient une chaîne qui représente la version de l’application hôte ou du serveur Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="6b2a8-130">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="6b2a8-p102">Si le complément de messagerie s’exécute sur le client de bureau Outlook ou sur Outlook pour iOS, la propriété `hostVersion` renvoie la version de l’application hôte, Outlook. Dans Outlook Web App, la propriété renvoie la version du serveur Exchange. Exemple : la chaîne `15.0.468.0`.</span><span class="sxs-lookup"><span data-stu-id="6b2a8-p102">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="6b2a8-134">Type :</span><span class="sxs-lookup"><span data-stu-id="6b2a8-134">Type:</span></span>

*   <span data-ttu-id="6b2a8-135">Chaîne</span><span class="sxs-lookup"><span data-stu-id="6b2a8-135">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6b2a8-136">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6b2a8-136">Requirements</span></span>

|<span data-ttu-id="6b2a8-137">Condition requise</span><span class="sxs-lookup"><span data-stu-id="6b2a8-137">Requirement</span></span>| <span data-ttu-id="6b2a8-138">Valeur</span><span class="sxs-lookup"><span data-stu-id="6b2a8-138">Value</span></span>|
|---|---|
|[<span data-ttu-id="6b2a8-139">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6b2a8-139">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6b2a8-140">1.0</span><span class="sxs-lookup"><span data-stu-id="6b2a8-140">1.0</span></span>|
|[<span data-ttu-id="6b2a8-141">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6b2a8-141">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6b2a8-142">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6b2a8-142">ReadItem</span></span>|
|[<span data-ttu-id="6b2a8-143">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6b2a8-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6b2a8-144">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="6b2a8-144">Compose or read</span></span>|

####  <a name="owaview-string"></a><span data-ttu-id="6b2a8-145">OWAView :String</span><span class="sxs-lookup"><span data-stu-id="6b2a8-145">OWAView :String</span></span>

<span data-ttu-id="6b2a8-146">Obtient une chaîne qui représente le mode d’affichage actuel dans Outlook Web App.</span><span class="sxs-lookup"><span data-stu-id="6b2a8-146">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="6b2a8-147">La chaîne renvoyée peut avoir une des valeurs suivantes : `OneColumn`, `TwoColumns`, ou `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="6b2a8-147">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="6b2a8-148">Si l’application hôte n’est pas Outlook Web App, l’accès à cette propriété retourne la valeur `undefined`.</span><span class="sxs-lookup"><span data-stu-id="6b2a8-148">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="6b2a8-149">Outlook Web App a trois modes d’affichage qui correspondent à la largeur de l’écran et de la fenêtre, ainsi qu’au nombre de colonnes pouvant être affichées :</span><span class="sxs-lookup"><span data-stu-id="6b2a8-149">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="6b2a8-p103">`OneColumn`, qui est affiché lorsque l’écran est étroit. Outlook Web App offre une mise en page à une colonne sur l’ensemble de l’écran d’un smartphone.</span><span class="sxs-lookup"><span data-stu-id="6b2a8-p103">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="6b2a8-p104">`TwoColumns`, qui est affiché lorsque l’écran est plus large. Outlook Web App utilise ce mode sur la plupart des tablettes.</span><span class="sxs-lookup"><span data-stu-id="6b2a8-p104">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="6b2a8-p105">`ThreeColumns`, qui est affiché lorsque l’écran est large. Par exemple, Outlook Web App utilise ce mode dans une fenêtre en mode plein écran sur un ordinateur de bureau.</span><span class="sxs-lookup"><span data-stu-id="6b2a8-p105">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="6b2a8-156">Type :</span><span class="sxs-lookup"><span data-stu-id="6b2a8-156">Type:</span></span>

*   <span data-ttu-id="6b2a8-157">Chaîne</span><span class="sxs-lookup"><span data-stu-id="6b2a8-157">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6b2a8-158">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6b2a8-158">Requirements</span></span>

|<span data-ttu-id="6b2a8-159">Condition requise</span><span class="sxs-lookup"><span data-stu-id="6b2a8-159">Requirement</span></span>| <span data-ttu-id="6b2a8-160">Valeur</span><span class="sxs-lookup"><span data-stu-id="6b2a8-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="6b2a8-161">Version minimale de l’ensemble des conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6b2a8-161">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6b2a8-162">1.0</span><span class="sxs-lookup"><span data-stu-id="6b2a8-162">1.0</span></span>|
|[<span data-ttu-id="6b2a8-163">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6b2a8-163">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6b2a8-164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="6b2a8-164">ReadItem</span></span>|
|[<span data-ttu-id="6b2a8-165">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6b2a8-165">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6b2a8-166">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="6b2a8-166">Compose or read</span></span>|