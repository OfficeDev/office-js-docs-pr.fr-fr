---
title: Office. Context. Mailbox. Diagnostics-ensemble de conditions requises
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 3bcc1004ddcee638965c15c7f320e2320a6faf06
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871898"
---
# <a name="diagnostics"></a><span data-ttu-id="47b4c-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="47b4c-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="47b4c-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="47b4c-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="47b4c-104">Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="47b4c-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="47b4c-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="47b4c-105">Requirements</span></span>

|<span data-ttu-id="47b4c-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="47b4c-106">Requirement</span></span>| <span data-ttu-id="47b4c-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="47b4c-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="47b4c-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="47b4c-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="47b4c-109">1.0</span><span class="sxs-lookup"><span data-stu-id="47b4c-109">1.0</span></span>|
|[<span data-ttu-id="47b4c-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="47b4c-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="47b4c-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="47b4c-111">ReadItem</span></span>|
|[<span data-ttu-id="47b4c-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="47b4c-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="47b4c-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="47b4c-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="47b4c-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="47b4c-114">Members and methods</span></span>

| <span data-ttu-id="47b4c-115">Membre</span><span class="sxs-lookup"><span data-stu-id="47b4c-115">Member</span></span> | <span data-ttu-id="47b4c-116">Type</span><span class="sxs-lookup"><span data-stu-id="47b4c-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="47b4c-117">Nom-d'hôte</span><span class="sxs-lookup"><span data-stu-id="47b4c-117">hostName</span></span>](#hostname-string) | <span data-ttu-id="47b4c-118">Member</span><span class="sxs-lookup"><span data-stu-id="47b4c-118">Member</span></span> |
| [<span data-ttu-id="47b4c-119">hostVersion</span><span class="sxs-lookup"><span data-stu-id="47b4c-119">hostVersion</span></span>](#hostversion-string) | <span data-ttu-id="47b4c-120">Member</span><span class="sxs-lookup"><span data-stu-id="47b4c-120">Member</span></span> |
| [<span data-ttu-id="47b4c-121">OWAView</span><span class="sxs-lookup"><span data-stu-id="47b4c-121">OWAView</span></span>](#owaview-string) | <span data-ttu-id="47b4c-122">Membre</span><span class="sxs-lookup"><span data-stu-id="47b4c-122">Member</span></span> |

### <a name="members"></a><span data-ttu-id="47b4c-123">Membres</span><span class="sxs-lookup"><span data-stu-id="47b4c-123">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="47b4c-124">hostName :String</span><span class="sxs-lookup"><span data-stu-id="47b4c-124">hostName :String</span></span>

<span data-ttu-id="47b4c-125">Obtient une chaîne qui représente le nom de l’application hôte.</span><span class="sxs-lookup"><span data-stu-id="47b4c-125">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="47b4c-126">Chaîne qui peut avoir l’une des valeurs suivantes : `Outlook`, `Mac Outlook`, `OutlookIOS` ou `OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="47b4c-126">A string that can be one of the following values: `Outlook`, `Mac Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="47b4c-127">Type</span><span class="sxs-lookup"><span data-stu-id="47b4c-127">Type</span></span>

*   <span data-ttu-id="47b4c-128">String</span><span class="sxs-lookup"><span data-stu-id="47b4c-128">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="47b4c-129">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="47b4c-129">Requirements</span></span>

|<span data-ttu-id="47b4c-130">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="47b4c-130">Requirement</span></span>| <span data-ttu-id="47b4c-131">Valeur</span><span class="sxs-lookup"><span data-stu-id="47b4c-131">Value</span></span>|
|---|---|
|[<span data-ttu-id="47b4c-132">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="47b4c-132">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="47b4c-133">1.0</span><span class="sxs-lookup"><span data-stu-id="47b4c-133">1.0</span></span>|
|[<span data-ttu-id="47b4c-134">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="47b4c-134">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="47b4c-135">ReadItem</span><span class="sxs-lookup"><span data-stu-id="47b4c-135">ReadItem</span></span>|
|[<span data-ttu-id="47b4c-136">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="47b4c-136">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="47b4c-137">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="47b4c-137">Compose or Read</span></span>|

####  <a name="hostversion-string"></a><span data-ttu-id="47b4c-138">hostVersion :String</span><span class="sxs-lookup"><span data-stu-id="47b4c-138">hostVersion :String</span></span>

<span data-ttu-id="47b4c-139">Obtient une chaîne qui représente la version de l’application hôte ou du serveur Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="47b4c-139">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="47b4c-p101">Si le complément de messagerie s’exécute sur le client de bureau Outlook ou sur Outlook pour iOS, la propriété `hostVersion` renvoie la version de l’application hôte, Outlook. Dans Outlook Web App, la propriété renvoie la version du serveur Exchange. Exemple : la chaîne `15.0.468.0`.</span><span class="sxs-lookup"><span data-stu-id="47b4c-p101">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="47b4c-143">Type</span><span class="sxs-lookup"><span data-stu-id="47b4c-143">Type</span></span>

*   <span data-ttu-id="47b4c-144">String</span><span class="sxs-lookup"><span data-stu-id="47b4c-144">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="47b4c-145">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="47b4c-145">Requirements</span></span>

|<span data-ttu-id="47b4c-146">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="47b4c-146">Requirement</span></span>| <span data-ttu-id="47b4c-147">Valeur</span><span class="sxs-lookup"><span data-stu-id="47b4c-147">Value</span></span>|
|---|---|
|[<span data-ttu-id="47b4c-148">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="47b4c-148">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="47b4c-149">1.0</span><span class="sxs-lookup"><span data-stu-id="47b4c-149">1.0</span></span>|
|[<span data-ttu-id="47b4c-150">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="47b4c-150">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="47b4c-151">ReadItem</span><span class="sxs-lookup"><span data-stu-id="47b4c-151">ReadItem</span></span>|
|[<span data-ttu-id="47b4c-152">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="47b4c-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="47b4c-153">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="47b4c-153">Compose or Read</span></span>|

####  <a name="owaview-string"></a><span data-ttu-id="47b4c-154">OWAView :String</span><span class="sxs-lookup"><span data-stu-id="47b4c-154">OWAView :String</span></span>

<span data-ttu-id="47b4c-155">Obtient une chaîne qui représente le mode d’affichage actuel dans Outlook Web App.</span><span class="sxs-lookup"><span data-stu-id="47b4c-155">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="47b4c-156">La chaîne renvoyée peut avoir une des valeurs suivantes : `OneColumn`, `TwoColumns`, ou `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="47b4c-156">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="47b4c-157">Si l’application hôte n’est pas Outlook Web App, l’accès à cette propriété génère la valeur `undefined`.</span><span class="sxs-lookup"><span data-stu-id="47b4c-157">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="47b4c-158">Outlook Web App a trois modes d’affichage qui correspondent à la largeur de l’écran et de la fenêtre, ainsi qu’au nombre de colonnes pouvant être affichées :</span><span class="sxs-lookup"><span data-stu-id="47b4c-158">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="47b4c-p102">`OneColumn`, qui est affiché lorsque l’écran est étroit. Outlook Web App offre une mise en page à une colonne sur l’ensemble de l’écran d’un smartphone.</span><span class="sxs-lookup"><span data-stu-id="47b4c-p102">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="47b4c-p103">`TwoColumns`, qui est affiché lorsque l’écran est plus large. Outlook Web App utilise ce mode sur la plupart des tablettes.</span><span class="sxs-lookup"><span data-stu-id="47b4c-p103">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="47b4c-p104">`ThreeColumns`, qui est affiché lorsque l’écran est large. Par exemple, Outlook Web App utilise ce mode dans une fenêtre en mode Plein écran sur un ordinateur de bureau.</span><span class="sxs-lookup"><span data-stu-id="47b4c-p104">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="47b4c-165">Type</span><span class="sxs-lookup"><span data-stu-id="47b4c-165">Type</span></span>

*   <span data-ttu-id="47b4c-166">String</span><span class="sxs-lookup"><span data-stu-id="47b4c-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="47b4c-167">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="47b4c-167">Requirements</span></span>

|<span data-ttu-id="47b4c-168">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="47b4c-168">Requirement</span></span>| <span data-ttu-id="47b4c-169">Valeur</span><span class="sxs-lookup"><span data-stu-id="47b4c-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="47b4c-170">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="47b4c-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="47b4c-171">1.0</span><span class="sxs-lookup"><span data-stu-id="47b4c-171">1.0</span></span>|
|[<span data-ttu-id="47b4c-172">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="47b4c-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="47b4c-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="47b4c-173">ReadItem</span></span>|
|[<span data-ttu-id="47b4c-174">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="47b4c-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="47b4c-175">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="47b4c-175">Compose or Read</span></span>|
