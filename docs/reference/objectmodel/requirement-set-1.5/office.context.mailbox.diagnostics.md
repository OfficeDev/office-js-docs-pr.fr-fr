---
title: Office. Context. Mailbox. Diagnostics-ensemble de conditions requises 1,5
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 7e554217831f2739ead3a0a90bd41b7d72e7b2d1
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450246"
---
# <a name="diagnostics"></a><span data-ttu-id="d4f52-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="d4f52-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="d4f52-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="d4f52-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="d4f52-104">Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="d4f52-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d4f52-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d4f52-105">Requirements</span></span>

|<span data-ttu-id="d4f52-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d4f52-106">Requirement</span></span>| <span data-ttu-id="d4f52-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="d4f52-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4f52-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d4f52-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4f52-109">1.0</span><span class="sxs-lookup"><span data-stu-id="d4f52-109">1.0</span></span>|
|[<span data-ttu-id="d4f52-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d4f52-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4f52-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d4f52-111">ReadItem</span></span>|
|[<span data-ttu-id="d4f52-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d4f52-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d4f52-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d4f52-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="d4f52-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="d4f52-114">Members and methods</span></span>

| <span data-ttu-id="d4f52-115">Membre</span><span class="sxs-lookup"><span data-stu-id="d4f52-115">Member</span></span> | <span data-ttu-id="d4f52-116">Type</span><span class="sxs-lookup"><span data-stu-id="d4f52-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="d4f52-117">Nom-d'hôte</span><span class="sxs-lookup"><span data-stu-id="d4f52-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="d4f52-118">Member</span><span class="sxs-lookup"><span data-stu-id="d4f52-118">Member</span></span> |
| [<span data-ttu-id="d4f52-119">hostVersion</span><span class="sxs-lookup"><span data-stu-id="d4f52-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="d4f52-120">Member</span><span class="sxs-lookup"><span data-stu-id="d4f52-120">Member</span></span> |
| [<span data-ttu-id="d4f52-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="d4f52-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="d4f52-122">Membre</span><span class="sxs-lookup"><span data-stu-id="d4f52-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="d4f52-123">Membres</span><span class="sxs-lookup"><span data-stu-id="d4f52-123">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="d4f52-124">hostName :String</span><span class="sxs-lookup"><span data-stu-id="d4f52-124">hostName :String</span></span>

<span data-ttu-id="d4f52-125">Obtient une chaîne qui représente le nom de l’application hôte.</span><span class="sxs-lookup"><span data-stu-id="d4f52-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="d4f52-126">Une chaîne qui peut avoir l’une des valeurs suivantes: `Outlook`, `OutlookIOS`ou`OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="d4f52-126">A string that can be one of the following values: `Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="d4f52-127">Type</span><span class="sxs-lookup"><span data-stu-id="d4f52-127">Type</span></span>

*   <span data-ttu-id="d4f52-128">String</span><span class="sxs-lookup"><span data-stu-id="d4f52-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d4f52-129">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d4f52-129">Requirements</span></span>

|<span data-ttu-id="d4f52-130">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d4f52-130">Requirement</span></span>| <span data-ttu-id="d4f52-131">Valeur</span><span class="sxs-lookup"><span data-stu-id="d4f52-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4f52-132">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d4f52-132">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4f52-133">1.0</span><span class="sxs-lookup"><span data-stu-id="d4f52-133">1.0</span></span>|
|[<span data-ttu-id="d4f52-134">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d4f52-134">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4f52-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d4f52-135">ReadItem</span></span>|
|[<span data-ttu-id="d4f52-136">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d4f52-136">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d4f52-137">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d4f52-137">Compose or Read</span></span>|

####  <a name="hostversion-string"></a><span data-ttu-id="d4f52-138">hostVersion :String</span><span class="sxs-lookup"><span data-stu-id="d4f52-138">hostVersion :String</span></span>

<span data-ttu-id="d4f52-139">Obtient une chaîne qui représente la version de l’application hôte ou du serveur Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="d4f52-139">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="d4f52-p101">Si le complément de messagerie s’exécute sur le client de bureau Outlook ou sur Outlook pour iOS, la propriété `hostVersion` renvoie la version de l’application hôte, Outlook. Dans Outlook Web App, la propriété renvoie la version du serveur Exchange. Exemple : la chaîne `15.0.468.0`.</span><span class="sxs-lookup"><span data-stu-id="d4f52-p101">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="d4f52-143">Type</span><span class="sxs-lookup"><span data-stu-id="d4f52-143">Type</span></span>

*   <span data-ttu-id="d4f52-144">String</span><span class="sxs-lookup"><span data-stu-id="d4f52-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d4f52-145">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d4f52-145">Requirements</span></span>

|<span data-ttu-id="d4f52-146">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d4f52-146">Requirement</span></span>| <span data-ttu-id="d4f52-147">Valeur</span><span class="sxs-lookup"><span data-stu-id="d4f52-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4f52-148">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d4f52-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4f52-149">1.0</span><span class="sxs-lookup"><span data-stu-id="d4f52-149">1.0</span></span>|
|[<span data-ttu-id="d4f52-150">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d4f52-150">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4f52-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d4f52-151">ReadItem</span></span>|
|[<span data-ttu-id="d4f52-152">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d4f52-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d4f52-153">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d4f52-153">Compose or Read</span></span>|

####  <a name="owaview-string"></a><span data-ttu-id="d4f52-154">OWAView :String</span><span class="sxs-lookup"><span data-stu-id="d4f52-154">OWAView :String</span></span>

<span data-ttu-id="d4f52-155">Obtient une chaîne qui représente le mode d’affichage actuel dans Outlook Web App.</span><span class="sxs-lookup"><span data-stu-id="d4f52-155">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="d4f52-156">La chaîne renvoyée peut avoir une des valeurs suivantes : `OneColumn`, `TwoColumns`, ou `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="d4f52-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="d4f52-157">Si l’application hôte n’est pas Outlook Web App, l’accès à cette propriété génère la valeur `undefined`.</span><span class="sxs-lookup"><span data-stu-id="d4f52-157">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="d4f52-158">Outlook Web App a trois modes d’affichage qui correspondent à la largeur de l’écran et de la fenêtre, ainsi qu’au nombre de colonnes pouvant être affichées :</span><span class="sxs-lookup"><span data-stu-id="d4f52-158">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="d4f52-p102">`OneColumn`, qui est affiché lorsque l’écran est étroit. Outlook Web App offre une mise en page à une colonne sur l’ensemble de l’écran d’un smartphone.</span><span class="sxs-lookup"><span data-stu-id="d4f52-p102">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="d4f52-p103">`TwoColumns`, qui est affiché lorsque l’écran est plus large. Outlook Web App utilise ce mode sur la plupart des tablettes.</span><span class="sxs-lookup"><span data-stu-id="d4f52-p103">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="d4f52-p104">`ThreeColumns`, qui est affiché lorsque l’écran est large. Par exemple, Outlook Web App utilise ce mode dans une fenêtre en mode Plein écran sur un ordinateur de bureau.</span><span class="sxs-lookup"><span data-stu-id="d4f52-p104">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="d4f52-165">Type</span><span class="sxs-lookup"><span data-stu-id="d4f52-165">Type</span></span>

*   <span data-ttu-id="d4f52-166">String</span><span class="sxs-lookup"><span data-stu-id="d4f52-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d4f52-167">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d4f52-167">Requirements</span></span>

|<span data-ttu-id="d4f52-168">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d4f52-168">Requirement</span></span>| <span data-ttu-id="d4f52-169">Valeur</span><span class="sxs-lookup"><span data-stu-id="d4f52-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4f52-170">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d4f52-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4f52-171">1.0</span><span class="sxs-lookup"><span data-stu-id="d4f52-171">1.0</span></span>|
|[<span data-ttu-id="d4f52-172">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d4f52-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d4f52-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d4f52-173">ReadItem</span></span>|
|[<span data-ttu-id="d4f52-174">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d4f52-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d4f52-175">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d4f52-175">Compose or Read</span></span>|
