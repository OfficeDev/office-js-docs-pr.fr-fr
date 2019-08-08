---
title: Office. Context. Mailbox. Diagnostics-ensemble de conditions requises
description: ''
ms.date: 08/05/2019
localization_priority: Normal
ms.openlocfilehash: 99848d3d0bd15726a54583210f94c3da035cd97f
ms.sourcegitcommit: dc78ee2a89fe3d4cd6f748be1eec9081c1077502
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2019
ms.locfileid: "36231290"
---
# <a name="diagnostics"></a><span data-ttu-id="9f6c7-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="9f6c7-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="9f6c7-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="9f6c7-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="9f6c7-104">Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="9f6c7-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9f6c7-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9f6c7-105">Requirements</span></span>

|<span data-ttu-id="9f6c7-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9f6c7-106">Requirement</span></span>| <span data-ttu-id="9f6c7-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="9f6c7-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="9f6c7-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9f6c7-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9f6c7-109">1.0</span><span class="sxs-lookup"><span data-stu-id="9f6c7-109">1.0</span></span>|
|[<span data-ttu-id="9f6c7-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9f6c7-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9f6c7-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9f6c7-111">ReadItem</span></span>|
|[<span data-ttu-id="9f6c7-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9f6c7-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9f6c7-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="9f6c7-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="9f6c7-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="9f6c7-114">Members and methods</span></span>

| <span data-ttu-id="9f6c7-115">Membre</span><span class="sxs-lookup"><span data-stu-id="9f6c7-115">Member</span></span> | <span data-ttu-id="9f6c7-116">Type</span><span class="sxs-lookup"><span data-stu-id="9f6c7-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="9f6c7-117">Nom-d’hôte</span><span class="sxs-lookup"><span data-stu-id="9f6c7-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="9f6c7-118">Member</span><span class="sxs-lookup"><span data-stu-id="9f6c7-118">Member</span></span> |
| [<span data-ttu-id="9f6c7-119">hostVersion</span><span class="sxs-lookup"><span data-stu-id="9f6c7-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="9f6c7-120">Member</span><span class="sxs-lookup"><span data-stu-id="9f6c7-120">Member</span></span> |
| [<span data-ttu-id="9f6c7-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="9f6c7-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="9f6c7-122">Membre</span><span class="sxs-lookup"><span data-stu-id="9f6c7-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="9f6c7-123">Membres</span><span class="sxs-lookup"><span data-stu-id="9f6c7-123">Members</span></span>

#### <a name="hostname-string"></a><span data-ttu-id="9f6c7-124">NomHôte: chaîne</span><span class="sxs-lookup"><span data-stu-id="9f6c7-124">hostName: String</span></span>

<span data-ttu-id="9f6c7-125">Obtient une chaîne qui représente le nom de l’application hôte.</span><span class="sxs-lookup"><span data-stu-id="9f6c7-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="9f6c7-126">Chaîne qui peut avoir l’une des valeurs suivantes : `Outlook`, `OutlookWebApp`, `OutlookIOS` ou `OutlookAndroid`.</span><span class="sxs-lookup"><span data-stu-id="9f6c7-126">A string that can be one of the following values: `Outlook`, `OutlookWebApp`, `OutlookIOS`, or `OutlookAndroid`.</span></span>

##### <a name="type"></a><span data-ttu-id="9f6c7-127">Type</span><span class="sxs-lookup"><span data-stu-id="9f6c7-127">Type</span></span>

*   <span data-ttu-id="9f6c7-128">String</span><span class="sxs-lookup"><span data-stu-id="9f6c7-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9f6c7-129">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9f6c7-129">Requirements</span></span>

|<span data-ttu-id="9f6c7-130">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9f6c7-130">Requirement</span></span>| <span data-ttu-id="9f6c7-131">Valeur</span><span class="sxs-lookup"><span data-stu-id="9f6c7-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="9f6c7-132">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9f6c7-132">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9f6c7-133">1.0</span><span class="sxs-lookup"><span data-stu-id="9f6c7-133">1.0</span></span>|
|[<span data-ttu-id="9f6c7-134">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9f6c7-134">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9f6c7-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9f6c7-135">ReadItem</span></span>|
|[<span data-ttu-id="9f6c7-136">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9f6c7-136">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9f6c7-137">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="9f6c7-137">Compose or Read</span></span>|

---
---

#### <a name="hostversion-string"></a><span data-ttu-id="9f6c7-138">hostVersion: chaîne</span><span class="sxs-lookup"><span data-stu-id="9f6c7-138">hostVersion: String</span></span>

<span data-ttu-id="9f6c7-139">Obtient une chaîne qui représente la version de l’application hôte ou du serveur Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="9f6c7-139">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="9f6c7-140">Si le complément de messagerie est en cours d’exécution sur le client de bureau Outlook ou sur `hostVersion` iOS, la propriété renvoie la version de l’application hôte, Outlook.</span><span class="sxs-lookup"><span data-stu-id="9f6c7-140">If the mail add-in is running on the Outlook desktop client or on iOS, the `hostVersion` property returns the version of the host application, Outlook.</span></span> <span data-ttu-id="9f6c7-141">Dans Outlook sur le Web, la propriété renvoie la version du serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="9f6c7-141">In Outlook on the web, the property returns the version of the Exchange Server.</span></span> <span data-ttu-id="9f6c7-142">La chaîne «15.0.468.0» est un exemple.</span><span class="sxs-lookup"><span data-stu-id="9f6c7-142">An example is the string "15.0.468.0".</span></span>

##### <a name="type"></a><span data-ttu-id="9f6c7-143">Type</span><span class="sxs-lookup"><span data-stu-id="9f6c7-143">Type</span></span>

*   <span data-ttu-id="9f6c7-144">String</span><span class="sxs-lookup"><span data-stu-id="9f6c7-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9f6c7-145">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9f6c7-145">Requirements</span></span>

|<span data-ttu-id="9f6c7-146">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9f6c7-146">Requirement</span></span>| <span data-ttu-id="9f6c7-147">Valeur</span><span class="sxs-lookup"><span data-stu-id="9f6c7-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="9f6c7-148">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9f6c7-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9f6c7-149">1.0</span><span class="sxs-lookup"><span data-stu-id="9f6c7-149">1.0</span></span>|
|[<span data-ttu-id="9f6c7-150">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9f6c7-150">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9f6c7-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9f6c7-151">ReadItem</span></span>|
|[<span data-ttu-id="9f6c7-152">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9f6c7-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9f6c7-153">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="9f6c7-153">Compose or Read</span></span>|

---
---

#### <a name="owaview-string"></a><span data-ttu-id="9f6c7-154">OWAView: chaîne</span><span class="sxs-lookup"><span data-stu-id="9f6c7-154">OWAView: String</span></span>

<span data-ttu-id="9f6c7-155">Obtient une valeur de type String qui représente l’affichage actuel d’Outlook sur le Web.</span><span class="sxs-lookup"><span data-stu-id="9f6c7-155">Gets a string that represents the current view of Outlook on the web.</span></span>

<span data-ttu-id="9f6c7-156">La chaîne renvoyée peut avoir une des valeurs suivantes : `OneColumn`, `TwoColumns`, ou `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="9f6c7-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="9f6c7-157">Si l’application hôte n’est pas Outlook sur le Web, l’accès à cette propriété génère `undefined`.</span><span class="sxs-lookup"><span data-stu-id="9f6c7-157">If the host application is not Outlook on the web, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="9f6c7-158">Outlook sur le Web possède trois modes d’affichage qui correspondent à la largeur de l’écran et de la fenêtre, ainsi qu’au nombre de colonnes pouvant être affichées:</span><span class="sxs-lookup"><span data-stu-id="9f6c7-158">Outlook on the web has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="9f6c7-159">`OneColumn`, qui est affiché lorsque l’écran est étroit.</span><span class="sxs-lookup"><span data-stu-id="9f6c7-159">`OneColumn`, which is displayed when the screen is narrow.</span></span> <span data-ttu-id="9f6c7-160">Outlook sur le Web utilise cette disposition sur une seule colonne sur la totalité de l’écran d’un smartphone.</span><span class="sxs-lookup"><span data-stu-id="9f6c7-160">Outlook on the web uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="9f6c7-161">`TwoColumns`, qui est affiché lorsque l’écran est plus large.</span><span class="sxs-lookup"><span data-stu-id="9f6c7-161">`TwoColumns`, which is displayed when the screen is wider.</span></span> <span data-ttu-id="9f6c7-162">Outlook sur le Web utilise cet affichage sur la plupart des tablettes.</span><span class="sxs-lookup"><span data-stu-id="9f6c7-162">Outlook on the web uses this view on most tablets.</span></span>
*   <span data-ttu-id="9f6c7-163">`ThreeColumns`, qui est affiché lorsque l’écran est large.</span><span class="sxs-lookup"><span data-stu-id="9f6c7-163">`ThreeColumns`, which is displayed when the screen is wide.</span></span> <span data-ttu-id="9f6c7-164">Par exemple, Outlook sur le Web utilise cet affichage dans une fenêtre plein écran sur un ordinateur de bureau.</span><span class="sxs-lookup"><span data-stu-id="9f6c7-164">For example, Outlook on the web uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="9f6c7-165">Type</span><span class="sxs-lookup"><span data-stu-id="9f6c7-165">Type</span></span>

*   <span data-ttu-id="9f6c7-166">String</span><span class="sxs-lookup"><span data-stu-id="9f6c7-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9f6c7-167">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9f6c7-167">Requirements</span></span>

|<span data-ttu-id="9f6c7-168">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9f6c7-168">Requirement</span></span>| <span data-ttu-id="9f6c7-169">Valeur</span><span class="sxs-lookup"><span data-stu-id="9f6c7-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="9f6c7-170">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9f6c7-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9f6c7-171">1.0</span><span class="sxs-lookup"><span data-stu-id="9f6c7-171">1.0</span></span>|
|[<span data-ttu-id="9f6c7-172">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9f6c7-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9f6c7-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9f6c7-173">ReadItem</span></span>|
|[<span data-ttu-id="9f6c7-174">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9f6c7-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9f6c7-175">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="9f6c7-175">Compose or Read</span></span>|
