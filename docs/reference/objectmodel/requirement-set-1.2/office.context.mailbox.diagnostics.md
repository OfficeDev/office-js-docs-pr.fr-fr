---
title: Office. Context. Mailbox. Diagnostics-ensemble de conditions requises 1,2
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: f2e613816884a5c1c00e5b96565d378434747e8e
ms.sourcegitcommit: dc78ee2a89fe3d4cd6f748be1eec9081c1077502
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2019
ms.locfileid: "36231269"
---
# <a name="diagnostics"></a><span data-ttu-id="b8967-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="b8967-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="b8967-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="b8967-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="b8967-104">Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="b8967-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b8967-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b8967-105">Requirements</span></span>

|<span data-ttu-id="b8967-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b8967-106">Requirement</span></span>| <span data-ttu-id="b8967-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="b8967-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="b8967-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b8967-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b8967-109">1.0</span><span class="sxs-lookup"><span data-stu-id="b8967-109">1.0</span></span>|
|[<span data-ttu-id="b8967-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b8967-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b8967-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b8967-111">ReadItem</span></span>|
|[<span data-ttu-id="b8967-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b8967-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b8967-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b8967-113">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="b8967-114">Members</span><span class="sxs-lookup"><span data-stu-id="b8967-114">Members</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="b8967-115">NomHôte: chaîne</span><span class="sxs-lookup"><span data-stu-id="b8967-115">hostName: String</span></span>

<span data-ttu-id="b8967-116">Obtient une chaîne qui représente le nom de l’application hôte.</span><span class="sxs-lookup"><span data-stu-id="b8967-116">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="b8967-117">Une chaîne qui peut avoir l’une des valeurs suivantes: `Outlook`, `OutlookIOS`ou`OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="b8967-117">A string that can be one of the following values: `Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="b8967-118">Type</span><span class="sxs-lookup"><span data-stu-id="b8967-118">Type</span></span>

*   <span data-ttu-id="b8967-119">String</span><span class="sxs-lookup"><span data-stu-id="b8967-119">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b8967-120">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b8967-120">Requirements</span></span>

|<span data-ttu-id="b8967-121">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b8967-121">Requirement</span></span>| <span data-ttu-id="b8967-122">Valeur</span><span class="sxs-lookup"><span data-stu-id="b8967-122">Value</span></span>|
|---|---|
|[<span data-ttu-id="b8967-123">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b8967-123">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b8967-124">1.0</span><span class="sxs-lookup"><span data-stu-id="b8967-124">1.0</span></span>|
|[<span data-ttu-id="b8967-125">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b8967-125">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b8967-126">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b8967-126">ReadItem</span></span>|
|[<span data-ttu-id="b8967-127">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b8967-127">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b8967-128">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b8967-128">Compose or Read</span></span>|

#### <a name="hostversion-string"></a><span data-ttu-id="b8967-129">hostVersion: chaîne</span><span class="sxs-lookup"><span data-stu-id="b8967-129">hostVersion: String</span></span>

<span data-ttu-id="b8967-130">Obtient une chaîne qui représente la version de l’application hôte ou du serveur Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="b8967-130">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="b8967-131">Si le complément de messagerie est en cours d’exécution sur le client de bureau Outlook ou `hostVersion` sur iOS, la propriété renvoie la version de l’application hôte, Outlook.</span><span class="sxs-lookup"><span data-stu-id="b8967-131">If the mail add-in is running on the Outlook desktop client or iOS, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="b8967-132">Dans Outlook sur le Web, la propriété renvoie la version du serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="b8967-132">In Outlook on the web, the property returns the version of the Exchange Server.</span></span> <span data-ttu-id="b8967-133">La chaîne «15.0.468.0» est un exemple.</span><span class="sxs-lookup"><span data-stu-id="b8967-133">An example is the string "15.0.468.0".</span></span>

##### <a name="type"></a><span data-ttu-id="b8967-134">Type</span><span class="sxs-lookup"><span data-stu-id="b8967-134">Type</span></span>

*   <span data-ttu-id="b8967-135">String</span><span class="sxs-lookup"><span data-stu-id="b8967-135">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b8967-136">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b8967-136">Requirements</span></span>

|<span data-ttu-id="b8967-137">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b8967-137">Requirement</span></span>| <span data-ttu-id="b8967-138">Valeur</span><span class="sxs-lookup"><span data-stu-id="b8967-138">Value</span></span>|
|---|---|
|[<span data-ttu-id="b8967-139">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b8967-139">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b8967-140">1.0</span><span class="sxs-lookup"><span data-stu-id="b8967-140">1.0</span></span>|
|[<span data-ttu-id="b8967-141">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b8967-141">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b8967-142">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b8967-142">ReadItem</span></span>|
|[<span data-ttu-id="b8967-143">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b8967-143">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b8967-144">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b8967-144">Compose or Read</span></span>|

#### <a name="owaview-string"></a><span data-ttu-id="b8967-145">OWAView: chaîne</span><span class="sxs-lookup"><span data-stu-id="b8967-145">OWAView: String</span></span>

<span data-ttu-id="b8967-146">Obtient une valeur de type String qui représente l’affichage actuel d’Outlook sur le Web.</span><span class="sxs-lookup"><span data-stu-id="b8967-146">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="b8967-147">La chaîne renvoyée peut avoir une des valeurs suivantes : `OneColumn`, `TwoColumns`, ou `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="b8967-147">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="b8967-148">Si l’application hôte n’est pas Outlook sur le Web, l’accès à cette propriété génère `undefined`.</span><span class="sxs-lookup"><span data-stu-id="b8967-148">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="b8967-149">Outlook sur le Web possède trois modes d’affichage qui correspondent à la largeur de l’écran et de la fenêtre, ainsi qu’au nombre de colonnes pouvant être affichées:</span><span class="sxs-lookup"><span data-stu-id="b8967-149">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="b8967-150">`OneColumn`, qui est affiché lorsque l’écran est étroit.</span><span class="sxs-lookup"><span data-stu-id="b8967-150">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="b8967-151">Outlook sur le Web utilise cette disposition sur une seule colonne sur la totalité de l’écran d’un smartphone.</span><span class="sxs-lookup"><span data-stu-id="b8967-151">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="b8967-152">`TwoColumns`, qui est affiché lorsque l’écran est plus large.</span><span class="sxs-lookup"><span data-stu-id="b8967-152">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="b8967-153">Outlook sur le Web utilise cet affichage sur la plupart des tablettes.</span><span class="sxs-lookup"><span data-stu-id="b8967-153">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="b8967-154">`ThreeColumns`, qui est affiché lorsque l’écran est large.</span><span class="sxs-lookup"><span data-stu-id="b8967-154">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="b8967-155">Par exemple, Outlook sur le Web utilise cet affichage dans une fenêtre plein écran sur un ordinateur de bureau.</span><span class="sxs-lookup"><span data-stu-id="b8967-155">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="b8967-156">Type</span><span class="sxs-lookup"><span data-stu-id="b8967-156">Type</span></span>

*   <span data-ttu-id="b8967-157">String</span><span class="sxs-lookup"><span data-stu-id="b8967-157">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b8967-158">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b8967-158">Requirements</span></span>

|<span data-ttu-id="b8967-159">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b8967-159">Requirement</span></span>| <span data-ttu-id="b8967-160">Valeur</span><span class="sxs-lookup"><span data-stu-id="b8967-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="b8967-161">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b8967-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b8967-162">1.0</span><span class="sxs-lookup"><span data-stu-id="b8967-162">1.0</span></span>|
|[<span data-ttu-id="b8967-163">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b8967-163">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b8967-164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b8967-164">ReadItem</span></span>|
|[<span data-ttu-id="b8967-165">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b8967-165">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b8967-166">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b8967-166">Compose or Read</span></span>|
