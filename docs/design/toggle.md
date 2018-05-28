---
title: Composant de bouton bascule dans Office UI Fabric
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 61bd251ac4d61922f228cd035e221a625890afee
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="toggle-component-in-office-ui-fabric"></a><span data-ttu-id="546d9-102">Composant de bouton bascule dans Office UI Fabric</span><span class="sxs-lookup"><span data-stu-id="546d9-102">Toggle component in Office UI Fabric</span></span>

<span data-ttu-id="546d9-p101">Les boutons bascules sont des commutateurs physiques qui activent ou d?sactivent des ?l?ments. Utilisez les boutons bascules pour pr?senter deux options qui s?excluent mutuellement (par exemple, on et off), lorsque le choix d?une option provoque une action imm?diate.</span><span class="sxs-lookup"><span data-stu-id="546d9-p101">Toggles represent a physical switch to turn things on or off. Use toggles to present two mutually exclusive options (for example, on or off), where choosing an option results in an immediate action.</span></span>
  
#### <a name="example-toggle-in-a-task-pane"></a><span data-ttu-id="546d9-105">Exemple : bouton bascule dans un volet Office</span><span class="sxs-lookup"><span data-stu-id="546d9-105">Example: Toggle in a task pane</span></span>

![Image illustrant le composant de bouton bascule](../images/overview-with-app-toggle.png)

## <a name="best-practices"></a><span data-ttu-id="546d9-107">Meilleures pratiques</span><span class="sxs-lookup"><span data-stu-id="546d9-107">Best practices</span></span>

|<span data-ttu-id="546d9-108">**? faire**</span><span class="sxs-lookup"><span data-stu-id="546d9-108">**Do**</span></span>|<span data-ttu-id="546d9-109">**? ne pas faire**</span><span class="sxs-lookup"><span data-stu-id="546d9-109">**Don't**</span></span>|
|:------------|:--------------|
|<span data-ttu-id="546d9-110">Utiliser les boutons bascule pour les param?tres binaires lorsque les modifications sont imm?diatement appliqu?es.</span><span class="sxs-lookup"><span data-stu-id="546d9-110">Use toggles for binary settings when changes are immediately applied.</span></span><br/><br/>![Exemple de bouton bascule ? faire](../images/toggle-do.png)<br/>|<span data-ttu-id="546d9-112">Ne pas utiliser de boutons bascule si les utilisateurs doivent effectuer une ?tape suppl?mentaire avant que les modifications prennent effet.</span><span class="sxs-lookup"><span data-stu-id="546d9-112">Don?t use toggles if users must perform an extra step before changes take effect.</span></span><br/><br/>![Exemple de bouton bascule ? ne pas faire](../images/toggle-dont.png)<br/>|
|<span data-ttu-id="546d9-p102">Remplacer les ?tiquettes **On** et **Off** uniquement s?il existe des ?tiquettes plus sp?cifiques ? utiliser pour un param?tre. Utiliser des ?tiquettes courtes (3 ? 4 caract?res) qui repr?sentent des oppos?s binaires.</span><span class="sxs-lookup"><span data-stu-id="546d9-p102">Only replace the **On** and **Off** labels if there are more specific labels to use for a setting. Use short (3-4 character) labels that represent binary opposites.</span></span>| |

## <a name="variants"></a><span data-ttu-id="546d9-116">Variantes</span><span class="sxs-lookup"><span data-stu-id="546d9-116">Variants</span></span>

|<span data-ttu-id="546d9-117">**Variation**</span><span class="sxs-lookup"><span data-stu-id="546d9-117">**Variation**</span></span>|<span data-ttu-id="546d9-118">**Description**</span><span class="sxs-lookup"><span data-stu-id="546d9-118">**Description**</span></span>|<span data-ttu-id="546d9-119">**Exemple**</span><span class="sxs-lookup"><span data-stu-id="546d9-119">**Example**</span></span>|
|:------------|:--------------|:----------|
|<span data-ttu-id="546d9-120">**Enabled and checked (Activ? et s?lectionn?)**</span><span class="sxs-lookup"><span data-stu-id="546d9-120">**Enabled and checked**</span></span>|<span data-ttu-id="546d9-121">? utiliser lorsque l??tat bascul? est actif.</span><span class="sxs-lookup"><span data-stu-id="546d9-121">Use when the toggled state is active.</span></span>|![Image Enabled and checked (Activ? et s?lectionn?)](../images/toggle-enabled-on.png)<br/>|
|<span data-ttu-id="546d9-123">**Enabled and unchecked (Activ? et d?s?lectionn?)**</span><span class="sxs-lookup"><span data-stu-id="546d9-123">**Enabled and unchecked**</span></span>|<span data-ttu-id="546d9-124">? utiliser lorsque l??tat bascul? est inactif.</span><span class="sxs-lookup"><span data-stu-id="546d9-124">Use when the toggled state is inactive.</span></span>|![Image Enabled and unchecked (Activ? et d?s?lectionn?)](../images/toggle-enabled-off.png)<br/>|
|<span data-ttu-id="546d9-126">**Disabled and checked (D?sactiv? et s?lectionn?)**</span><span class="sxs-lookup"><span data-stu-id="546d9-126">**Disabled and checked**</span></span>|<span data-ttu-id="546d9-127">? utiliser lorsque l??tat actif ne peut pas ?tre modifi?.</span><span class="sxs-lookup"><span data-stu-id="546d9-127">Use when the active state cannot be changed.</span></span>|![Image Disabled and checked (D?sactiv? et s?lectionn?)](../images/toggle-disabled-on.png)<br/>|
|<span data-ttu-id="546d9-129">**Disabled and unchecked (D?sactiv? et d?s?lectionn?)**</span><span class="sxs-lookup"><span data-stu-id="546d9-129">**Disabled and unchecked**</span></span>|<span data-ttu-id="546d9-130">? utiliser lorsque l??tat inactif ne peut pas ?tre modifi?.</span><span class="sxs-lookup"><span data-stu-id="546d9-130">Use when the inactive state cannot be changed.</span></span>|![Image Disabled and unchecked (D?sactiv? et d?s?lectionn?)](../images/toggle-disabled-off.png)<br/>|

## <a name="implementation"></a><span data-ttu-id="546d9-132">Impl?mentation</span><span class="sxs-lookup"><span data-stu-id="546d9-132">Implementation</span></span>

<span data-ttu-id="546d9-133">Pour plus d?informations, reportez-vous ? [Bouton bascule](https://dev.office.com/fabric#/components/toggle) et [D?marrer avec un exemple de code Fabric React](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).</span><span class="sxs-lookup"><span data-stu-id="546d9-133">For details, see [Toggle](https://dev.office.com/fabric#/components/toggle) and [Getting started with Fabric React code sample](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).</span></span>

## <a name="see-also"></a><span data-ttu-id="546d9-134">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="546d9-134">See also</span></span>

- [<span data-ttu-id="546d9-135">Mod?les de conception UX</span><span class="sxs-lookup"><span data-stu-id="546d9-135">UX Design Patterns</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [<span data-ttu-id="546d9-136">Office UI Fabric dans des compl?ments Office</span><span class="sxs-lookup"><span data-stu-id="546d9-136">Office UI Fabric in Office Add-ins</span></span>](office-ui-fabric.md)
