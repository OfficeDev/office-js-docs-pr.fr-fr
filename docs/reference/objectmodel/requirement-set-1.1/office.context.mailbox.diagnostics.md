---
title: Office.context.mailbox.diagnostics – ensemble de conditions requises 1.1
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 32c4d7389555bf6dab7cff3822abb4a4edbebe80
ms.sourcegitcommit: dc78ee2a89fe3d4cd6f748be1eec9081c1077502
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2019
ms.locfileid: "36231297"
---
# <a name="diagnostics"></a><span data-ttu-id="991bf-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="991bf-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="991bf-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="991bf-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="991bf-104">Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="991bf-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="991bf-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="991bf-105">Requirements</span></span>

|<span data-ttu-id="991bf-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="991bf-106">Requirement</span></span>| <span data-ttu-id="991bf-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="991bf-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="991bf-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="991bf-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="991bf-109">1.0</span><span class="sxs-lookup"><span data-stu-id="991bf-109">1.0</span></span>|
|[<span data-ttu-id="991bf-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="991bf-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="991bf-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="991bf-111">ReadItem</span></span>|
|[<span data-ttu-id="991bf-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="991bf-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="991bf-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="991bf-113">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="991bf-114">Members</span><span class="sxs-lookup"><span data-stu-id="991bf-114">Members</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="991bf-115">NomHôte: chaîne</span><span class="sxs-lookup"><span data-stu-id="991bf-115">hostName: String</span></span>

<span data-ttu-id="991bf-116">Obtient une chaîne qui représente le nom de l’application hôte.</span><span class="sxs-lookup"><span data-stu-id="991bf-116">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="991bf-117">Une chaîne qui peut avoir l’une des valeurs suivantes: `Outlook`, `OutlookIOS`ou`OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="991bf-117">A string that can be one of the following values: `Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="991bf-118">Type</span><span class="sxs-lookup"><span data-stu-id="991bf-118">Type</span></span>

*   <span data-ttu-id="991bf-119">String</span><span class="sxs-lookup"><span data-stu-id="991bf-119">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="991bf-120">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="991bf-120">Requirements</span></span>

|<span data-ttu-id="991bf-121">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="991bf-121">Requirement</span></span>| <span data-ttu-id="991bf-122">Valeur</span><span class="sxs-lookup"><span data-stu-id="991bf-122">Value</span></span>|
|---|---|
|[<span data-ttu-id="991bf-123">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="991bf-123">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="991bf-124">1.0</span><span class="sxs-lookup"><span data-stu-id="991bf-124">1.0</span></span>|
|[<span data-ttu-id="991bf-125">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="991bf-125">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="991bf-126">ReadItem</span><span class="sxs-lookup"><span data-stu-id="991bf-126">ReadItem</span></span>|
|[<span data-ttu-id="991bf-127">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="991bf-127">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="991bf-128">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="991bf-128">Compose or Read</span></span>|

#### <a name="hostversion-string"></a><span data-ttu-id="991bf-129">hostVersion: chaîne</span><span class="sxs-lookup"><span data-stu-id="991bf-129">hostVersion: String</span></span>

<span data-ttu-id="991bf-130">Obtient une chaîne qui représente la version de l’application hôte ou du serveur Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="991bf-130">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="991bf-131">Si le complément de messagerie est en cours d’exécution sur le client de bureau Outlook ou `hostVersion` sur iOS, la propriété renvoie la version de l’application hôte, Outlook.</span><span class="sxs-lookup"><span data-stu-id="991bf-131">If the mail add-in is running on the Outlook desktop client or iOS, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="991bf-132">Dans Outlook sur le Web, la propriété renvoie la version du serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="991bf-132">In Outlook on the web, the property returns the version of the Exchange Server.</span></span> <span data-ttu-id="991bf-133">La chaîne «15.0.468.0» est un exemple.</span><span class="sxs-lookup"><span data-stu-id="991bf-133">An example is the string "15.0.468.0".</span></span>

##### <a name="type"></a><span data-ttu-id="991bf-134">Type</span><span class="sxs-lookup"><span data-stu-id="991bf-134">Type</span></span>

*   <span data-ttu-id="991bf-135">String</span><span class="sxs-lookup"><span data-stu-id="991bf-135">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="991bf-136">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="991bf-136">Requirements</span></span>

|<span data-ttu-id="991bf-137">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="991bf-137">Requirement</span></span>| <span data-ttu-id="991bf-138">Valeur</span><span class="sxs-lookup"><span data-stu-id="991bf-138">Value</span></span>|
|---|---|
|[<span data-ttu-id="991bf-139">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="991bf-139">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="991bf-140">1.0</span><span class="sxs-lookup"><span data-stu-id="991bf-140">1.0</span></span>|
|[<span data-ttu-id="991bf-141">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="991bf-141">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="991bf-142">ReadItem</span><span class="sxs-lookup"><span data-stu-id="991bf-142">ReadItem</span></span>|
|[<span data-ttu-id="991bf-143">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="991bf-143">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="991bf-144">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="991bf-144">Compose or Read</span></span>|

#### <a name="owaview-string"></a><span data-ttu-id="991bf-145">OWAView: chaîne</span><span class="sxs-lookup"><span data-stu-id="991bf-145">OWAView: String</span></span>

<span data-ttu-id="991bf-146">Obtient une valeur de type String qui représente l’affichage actuel d’Outlook sur le Web.</span><span class="sxs-lookup"><span data-stu-id="991bf-146">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="991bf-147">La chaîne renvoyée peut avoir une des valeurs suivantes : `OneColumn`, `TwoColumns`, ou `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="991bf-147">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="991bf-148">Si l’application hôte n’est pas Outlook sur le Web, l’accès à cette propriété génère `undefined`.</span><span class="sxs-lookup"><span data-stu-id="991bf-148">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="991bf-149">Outlook sur le Web possède trois modes d’affichage qui correspondent à la largeur de l’écran et de la fenêtre, ainsi qu’au nombre de colonnes pouvant être affichées:</span><span class="sxs-lookup"><span data-stu-id="991bf-149">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="991bf-150">`OneColumn`, qui est affiché lorsque l’écran est étroit.</span><span class="sxs-lookup"><span data-stu-id="991bf-150">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="991bf-151">Outlook sur le Web utilise cette disposition sur une seule colonne sur la totalité de l’écran d’un smartphone.</span><span class="sxs-lookup"><span data-stu-id="991bf-151">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="991bf-152">`TwoColumns`, qui est affiché lorsque l’écran est plus large.</span><span class="sxs-lookup"><span data-stu-id="991bf-152">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="991bf-153">Outlook sur le Web utilise cet affichage sur la plupart des tablettes.</span><span class="sxs-lookup"><span data-stu-id="991bf-153">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="991bf-154">`ThreeColumns`, qui est affiché lorsque l’écran est large.</span><span class="sxs-lookup"><span data-stu-id="991bf-154">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="991bf-155">Par exemple, Outlook sur le Web utilise cet affichage dans une fenêtre plein écran sur un ordinateur de bureau.</span><span class="sxs-lookup"><span data-stu-id="991bf-155">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="991bf-156">Type</span><span class="sxs-lookup"><span data-stu-id="991bf-156">Type</span></span>

*   <span data-ttu-id="991bf-157">String</span><span class="sxs-lookup"><span data-stu-id="991bf-157">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="991bf-158">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="991bf-158">Requirements</span></span>

|<span data-ttu-id="991bf-159">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="991bf-159">Requirement</span></span>| <span data-ttu-id="991bf-160">Valeur</span><span class="sxs-lookup"><span data-stu-id="991bf-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="991bf-161">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="991bf-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="991bf-162">1.0</span><span class="sxs-lookup"><span data-stu-id="991bf-162">1.0</span></span>|
|[<span data-ttu-id="991bf-163">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="991bf-163">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="991bf-164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="991bf-164">ReadItem</span></span>|
|[<span data-ttu-id="991bf-165">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="991bf-165">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="991bf-166">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="991bf-166">Compose or Read</span></span>|
