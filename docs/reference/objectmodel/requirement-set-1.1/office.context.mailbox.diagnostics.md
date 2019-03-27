---
title: Office.context.mailbox.diagnostics – ensemble de conditions requises 1.1
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 8ce3242a9e007e9aa686ddd098ab23ee922276e4
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871191"
---
# <a name="diagnostics"></a><span data-ttu-id="973fb-102">diagnostics</span><span class="sxs-lookup"><span data-stu-id="973fb-102">diagnostics</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a><span data-ttu-id="973fb-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span><span class="sxs-lookup"><span data-stu-id="973fb-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics</span></span>

<span data-ttu-id="973fb-104">Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="973fb-104">Provides diagnostic information to an Outlook add-in.</span></span>

##### <a name="requirements"></a><span data-ttu-id="973fb-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="973fb-105">Requirements</span></span>

|<span data-ttu-id="973fb-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="973fb-106">Requirement</span></span>| <span data-ttu-id="973fb-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="973fb-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="973fb-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="973fb-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="973fb-109">1.0</span><span class="sxs-lookup"><span data-stu-id="973fb-109">1.0</span></span>|
|[<span data-ttu-id="973fb-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="973fb-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="973fb-111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="973fb-111">ReadItem</span></span>|
|[<span data-ttu-id="973fb-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="973fb-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="973fb-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="973fb-113">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="973fb-114">Membres</span><span class="sxs-lookup"><span data-stu-id="973fb-114">Members</span></span>

####  <a name="hostname-string"></a><span data-ttu-id="973fb-115">hostName :String</span><span class="sxs-lookup"><span data-stu-id="973fb-115">hostName :String</span></span>

<span data-ttu-id="973fb-116">Obtient une chaîne qui représente le nom de l’application hôte.</span><span class="sxs-lookup"><span data-stu-id="973fb-116">Gets a string that represents the name of the host application.</span></span>

<span data-ttu-id="973fb-117">Une chaîne qui peut avoir l’une des valeurs suivantes: `Outlook`, `OutlookIOS`ou`OutlookWebApp`.</span><span class="sxs-lookup"><span data-stu-id="973fb-117">A string that can be one of the following values: `Outlook`, `OutlookIOS`, or `OutlookWebApp`.</span></span>

##### <a name="type"></a><span data-ttu-id="973fb-118">Type</span><span class="sxs-lookup"><span data-stu-id="973fb-118">Type</span></span>

*   <span data-ttu-id="973fb-119">String</span><span class="sxs-lookup"><span data-stu-id="973fb-119">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="973fb-120">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="973fb-120">Requirements</span></span>

|<span data-ttu-id="973fb-121">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="973fb-121">Requirement</span></span>| <span data-ttu-id="973fb-122">Valeur</span><span class="sxs-lookup"><span data-stu-id="973fb-122">Value</span></span>|
|---|---|
|[<span data-ttu-id="973fb-123">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="973fb-123">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="973fb-124">1.0</span><span class="sxs-lookup"><span data-stu-id="973fb-124">1.0</span></span>|
|[<span data-ttu-id="973fb-125">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="973fb-125">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="973fb-126">ReadItem</span><span class="sxs-lookup"><span data-stu-id="973fb-126">ReadItem</span></span>|
|[<span data-ttu-id="973fb-127">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="973fb-127">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="973fb-128">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="973fb-128">Compose or Read</span></span>|

####  <a name="hostversion-string"></a><span data-ttu-id="973fb-129">hostVersion :String</span><span class="sxs-lookup"><span data-stu-id="973fb-129">hostVersion :String</span></span>

<span data-ttu-id="973fb-130">Obtient une chaîne qui représente la version de l’application hôte ou du serveur Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="973fb-130">Gets a string that represents the version of either the host application or the Exchange Server.</span></span>

<span data-ttu-id="973fb-p101">Si le complément de messagerie s’exécute sur le client de bureau Outlook ou sur Outlook pour iOS, la propriété `hostVersion` renvoie la version de l’application hôte, Outlook. Dans Outlook Web App, la propriété renvoie la version du serveur Exchange. Exemple : la chaîne `15.0.468.0`.</span><span class="sxs-lookup"><span data-stu-id="973fb-p101">If the mail add-in is running on the Outlook desktop client or Outlook for iOS, the `hostVersion` property returns the version of the host application, Outlook. In Outlook Web App, the property returns the version of the Exchange Server. An example is the string `15.0.468.0`.</span></span>

##### <a name="type"></a><span data-ttu-id="973fb-134">Type</span><span class="sxs-lookup"><span data-stu-id="973fb-134">Type</span></span>

*   <span data-ttu-id="973fb-135">String</span><span class="sxs-lookup"><span data-stu-id="973fb-135">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="973fb-136">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="973fb-136">Requirements</span></span>

|<span data-ttu-id="973fb-137">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="973fb-137">Requirement</span></span>| <span data-ttu-id="973fb-138">Valeur</span><span class="sxs-lookup"><span data-stu-id="973fb-138">Value</span></span>|
|---|---|
|[<span data-ttu-id="973fb-139">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="973fb-139">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="973fb-140">1.0</span><span class="sxs-lookup"><span data-stu-id="973fb-140">1.0</span></span>|
|[<span data-ttu-id="973fb-141">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="973fb-141">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="973fb-142">ReadItem</span><span class="sxs-lookup"><span data-stu-id="973fb-142">ReadItem</span></span>|
|[<span data-ttu-id="973fb-143">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="973fb-143">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="973fb-144">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="973fb-144">Compose or Read</span></span>|

####  <a name="owaview-string"></a><span data-ttu-id="973fb-145">OWAView :String</span><span class="sxs-lookup"><span data-stu-id="973fb-145">OWAView :String</span></span>

<span data-ttu-id="973fb-146">Obtient une chaîne qui représente le mode d’affichage actuel dans Outlook Web App.</span><span class="sxs-lookup"><span data-stu-id="973fb-146">Gets a string that represents the current view of Outlook Web App.</span></span>

<span data-ttu-id="973fb-147">La chaîne renvoyée peut avoir une des valeurs suivantes : `OneColumn`, `TwoColumns`, ou `ThreeColumns`.</span><span class="sxs-lookup"><span data-stu-id="973fb-147">The returned string can be one of the following values: `OneColumn`, `TwoColumns`, or `ThreeColumns`.</span></span>

<span data-ttu-id="973fb-148">Si l’application hôte n’est pas Outlook Web App, l’accès à cette propriété génère la valeur `undefined`.</span><span class="sxs-lookup"><span data-stu-id="973fb-148">If the host application is not Outlook Web App, then accessing this property results in `undefined`.</span></span>

<span data-ttu-id="973fb-149">Outlook Web App a trois modes d’affichage qui correspondent à la largeur de l’écran et de la fenêtre, ainsi qu’au nombre de colonnes pouvant être affichées :</span><span class="sxs-lookup"><span data-stu-id="973fb-149">Outlook Web App has three views that correspond to the width of the screen and the window, and the number of columns that can be displayed:</span></span>

*   <span data-ttu-id="973fb-p102">`OneColumn`, qui est affiché lorsque l’écran est étroit. Outlook Web App offre une mise en page à une colonne sur l’ensemble de l’écran d’un smartphone.</span><span class="sxs-lookup"><span data-stu-id="973fb-p102">`OneColumn`, which is displayed when the screen is narrow. Outlook Web App uses this single-column layout on the entire screen of a smartphone.</span></span>
*   <span data-ttu-id="973fb-p103">`TwoColumns`, qui est affiché lorsque l’écran est plus large. Outlook Web App utilise ce mode sur la plupart des tablettes.</span><span class="sxs-lookup"><span data-stu-id="973fb-p103">`TwoColumns`, which is displayed when the screen is wider. Outlook Web App uses this view on most tablets.</span></span>
*   <span data-ttu-id="973fb-p104">`ThreeColumns`, qui est affiché lorsque l’écran est large. Par exemple, Outlook Web App utilise ce mode dans une fenêtre en mode Plein écran sur un ordinateur de bureau.</span><span class="sxs-lookup"><span data-stu-id="973fb-p104">`ThreeColumns`, which is displayed when the screen is wide. For example, Outlook Web App uses this view in a full screen window on a desktop computer.</span></span>

##### <a name="type"></a><span data-ttu-id="973fb-156">Type</span><span class="sxs-lookup"><span data-stu-id="973fb-156">Type</span></span>

*   <span data-ttu-id="973fb-157">String</span><span class="sxs-lookup"><span data-stu-id="973fb-157">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="973fb-158">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="973fb-158">Requirements</span></span>

|<span data-ttu-id="973fb-159">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="973fb-159">Requirement</span></span>| <span data-ttu-id="973fb-160">Valeur</span><span class="sxs-lookup"><span data-stu-id="973fb-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="973fb-161">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="973fb-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="973fb-162">1.0</span><span class="sxs-lookup"><span data-stu-id="973fb-162">1.0</span></span>|
|[<span data-ttu-id="973fb-163">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="973fb-163">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="973fb-164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="973fb-164">ReadItem</span></span>|
|[<span data-ttu-id="973fb-165">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="973fb-165">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="973fb-166">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="973fb-166">Compose or Read</span></span>|
