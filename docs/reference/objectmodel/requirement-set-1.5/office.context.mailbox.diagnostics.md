---
title: Office. Context. Mailbox. Diagnostics-ensemble de conditions requises 1,5
description: ''
ms.date: 04/24/2019
localization_priority: Normal
ms.openlocfilehash: 9ecbf4382f10b86ecdea41706211094029be09d2
ms.sourcegitcommit: dc78ee2a89fe3d4cd6f748be1eec9081c1077502
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2019
ms.locfileid: "36231255"
---
# <a name="diagnostics"></a><span data-ttu-id="1b1b3-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="1b1b3-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="1b1b3-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="1b1b3-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="1b1b3-104">Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="1b1b3-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1b1b3-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1b1b3-105">Requirements</span></span>

|<span data-ttu-id="1b1b3-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1b1b3-106">Requirement</span></span>| <span data-ttu-id="1b1b3-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="1b1b3-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b1b3-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1b1b3-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b1b3-109">1.0</span><span class="sxs-lookup"><span data-stu-id="1b1b3-109">1.0</span></span>|
|[<span data-ttu-id="1b1b3-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1b1b3-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b1b3-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b1b3-111">ReadItem</span></span>|
|[<span data-ttu-id="1b1b3-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1b1b3-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b1b3-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="1b1b3-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="1b1b3-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="1b1b3-114">Members and methods</span></span>

| <span data-ttu-id="1b1b3-115">Membre</span><span class="sxs-lookup"><span data-stu-id="1b1b3-115">Member</span></span> | <span data-ttu-id="1b1b3-116">Type</span><span class="sxs-lookup"><span data-stu-id="1b1b3-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="1b1b3-117">Nom-d’hôte</span><span class="sxs-lookup"><span data-stu-id="1b1b3-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="1b1b3-118">Member</span><span class="sxs-lookup"><span data-stu-id="1b1b3-118">Member</span></span> |
| [<span data-ttu-id="1b1b3-119">hostVersion</span><span class="sxs-lookup"><span data-stu-id="1b1b3-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="1b1b3-120">Member</span><span class="sxs-lookup"><span data-stu-id="1b1b3-120">Member</span></span> |
| [<span data-ttu-id="1b1b3-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="1b1b3-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="1b1b3-122">Membre</span><span class="sxs-lookup"><span data-stu-id="1b1b3-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="1b1b3-123">Membres</span><span class="sxs-lookup"><span data-stu-id="1b1b3-123">Members</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="1b1b3-124">NomHôte: chaîne</span><span class="sxs-lookup"><span data-stu-id="1b1b3-124">hostName: String</span></span>

<span data-ttu-id="1b1b3-125">Obtient une chaîne qui représente le nom de l’application hôte.</span><span class="sxs-lookup"><span data-stu-id="1b1b3-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="1b1b3-126">Une chaîne qui peut avoir l’une des valeurs suivantes: `Outlook`, `OutlookIOS`ou`OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="1b1b3-126">A string that can be one of the following values: `Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="1b1b3-127">Type</span><span class="sxs-lookup"><span data-stu-id="1b1b3-127">Type</span></span>

*   <span data-ttu-id="1b1b3-128">String</span><span class="sxs-lookup"><span data-stu-id="1b1b3-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1b1b3-129">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1b1b3-129">Requirements</span></span>

|<span data-ttu-id="1b1b3-130">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1b1b3-130">Requirement</span></span>| <span data-ttu-id="1b1b3-131">Valeur</span><span class="sxs-lookup"><span data-stu-id="1b1b3-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b1b3-132">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1b1b3-132">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b1b3-133">1.0</span><span class="sxs-lookup"><span data-stu-id="1b1b3-133">1.0</span></span>|
|[<span data-ttu-id="1b1b3-134">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1b1b3-134">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b1b3-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b1b3-135">ReadItem</span></span>|
|[<span data-ttu-id="1b1b3-136">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1b1b3-136">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b1b3-137">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="1b1b3-137">Compose or Read</span></span>|

#### <a name="hostversion-string"></a><span data-ttu-id="1b1b3-138">hostVersion: chaîne</span><span class="sxs-lookup"><span data-stu-id="1b1b3-138">hostVersion: String</span></span>

<span data-ttu-id="1b1b3-139">Obtient une chaîne qui représente la version de l’application hôte ou du serveur Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="1b1b3-139">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="1b1b3-140">Si le complément de messagerie est en cours d’exécution sur le client de bureau Outlook ou `hostVersion` sur iOS, la propriété renvoie la version de l’application hôte, Outlook.</span><span class="sxs-lookup"><span data-stu-id="1b1b3-140">If the mail add-in is running on the Outlook desktop client or iOS, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="1b1b3-141">Dans Outlook sur le Web, la propriété renvoie la version du serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="1b1b3-141">In Outlook on the web, the property returns the version of the Exchange Server.</span></span> <span data-ttu-id="1b1b3-142">La chaîne «15.0.468.0» est un exemple.</span><span class="sxs-lookup"><span data-stu-id="1b1b3-142">An example is the string "15.0.468.0".</span></span>

##### <a name="type"></a><span data-ttu-id="1b1b3-143">Type</span><span class="sxs-lookup"><span data-stu-id="1b1b3-143">Type</span></span>

*   <span data-ttu-id="1b1b3-144">String</span><span class="sxs-lookup"><span data-stu-id="1b1b3-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1b1b3-145">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1b1b3-145">Requirements</span></span>

|<span data-ttu-id="1b1b3-146">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1b1b3-146">Requirement</span></span>| <span data-ttu-id="1b1b3-147">Valeur</span><span class="sxs-lookup"><span data-stu-id="1b1b3-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b1b3-148">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1b1b3-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b1b3-149">1.0</span><span class="sxs-lookup"><span data-stu-id="1b1b3-149">1.0</span></span>|
|[<span data-ttu-id="1b1b3-150">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1b1b3-150">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b1b3-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b1b3-151">ReadItem</span></span>|
|[<span data-ttu-id="1b1b3-152">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1b1b3-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b1b3-153">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="1b1b3-153">Compose or Read</span></span>|

#### <a name="owaview-string"></a><span data-ttu-id="1b1b3-154">OWAView: chaîne</span><span class="sxs-lookup"><span data-stu-id="1b1b3-154">OWAView: String</span></span>

<span data-ttu-id="1b1b3-155">Obtient une valeur de type String qui représente l’affichage actuel d’Outlook sur le Web.</span><span class="sxs-lookup"><span data-stu-id="1b1b3-155">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="1b1b3-156">La chaîne renvoyée peut avoir une des valeurs suivantes : `OneColumn`, `TwoColumns`, ou `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="1b1b3-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="1b1b3-157">Si l’application hôte n’est pas Outlook sur le Web, l’accès à cette propriété génère `undefined`.</span><span class="sxs-lookup"><span data-stu-id="1b1b3-157">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="1b1b3-158">Outlook sur le Web possède trois modes d’affichage qui correspondent à la largeur de l’écran et de la fenêtre, ainsi qu’au nombre de colonnes pouvant être affichées:</span><span class="sxs-lookup"><span data-stu-id="1b1b3-158">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="1b1b3-159">`OneColumn`, qui est affiché lorsque l’écran est étroit.</span><span class="sxs-lookup"><span data-stu-id="1b1b3-159">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="1b1b3-160">Outlook sur le Web utilise cette disposition sur une seule colonne sur la totalité de l’écran d’un smartphone.</span><span class="sxs-lookup"><span data-stu-id="1b1b3-160">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="1b1b3-161">`TwoColumns`, qui est affiché lorsque l’écran est plus large.</span><span class="sxs-lookup"><span data-stu-id="1b1b3-161">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="1b1b3-162">Outlook sur le Web utilise cet affichage sur la plupart des tablettes.</span><span class="sxs-lookup"><span data-stu-id="1b1b3-162">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="1b1b3-163">`ThreeColumns`, qui est affiché lorsque l’écran est large.</span><span class="sxs-lookup"><span data-stu-id="1b1b3-163">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="1b1b3-164">Par exemple, Outlook sur le Web utilise cet affichage dans une fenêtre plein écran sur un ordinateur de bureau.</span><span class="sxs-lookup"><span data-stu-id="1b1b3-164">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="1b1b3-165">Type</span><span class="sxs-lookup"><span data-stu-id="1b1b3-165">Type</span></span>

*   <span data-ttu-id="1b1b3-166">String</span><span class="sxs-lookup"><span data-stu-id="1b1b3-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1b1b3-167">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1b1b3-167">Requirements</span></span>

|<span data-ttu-id="1b1b3-168">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1b1b3-168">Requirement</span></span>| <span data-ttu-id="1b1b3-169">Valeur</span><span class="sxs-lookup"><span data-stu-id="1b1b3-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b1b3-170">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1b1b3-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1b1b3-171">1.0</span><span class="sxs-lookup"><span data-stu-id="1b1b3-171">1.0</span></span>|
|[<span data-ttu-id="1b1b3-172">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1b1b3-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1b1b3-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1b1b3-173">ReadItem</span></span>|
|[<span data-ttu-id="1b1b3-174">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1b1b3-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1b1b3-175">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="1b1b3-175">Compose or Read</span></span>|
