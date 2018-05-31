---
title: Composant de bouton bascule dans Office UI Fabric
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 61bd251ac4d61922f228cd035e221a625890afee
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437233"
---
# <a name="toggle-component-in-office-ui-fabric"></a><span data-ttu-id="3009d-102">Composant de bouton bascule dans Office UI Fabric</span><span class="sxs-lookup"><span data-stu-id="3009d-102">Toggle component in Office UI Fabric</span></span>

<span data-ttu-id="3009d-p101">Les boutons bascules sont des commutateurs physiques qui activent ou désactivent des éléments. Utilisez les boutons bascules pour présenter deux options qui s’excluent mutuellement (par exemple, on et off), lorsque le choix d’une option provoque une action immédiate.</span><span class="sxs-lookup"><span data-stu-id="3009d-p101">Toggles represent a physical switch to turn things on or off. Use toggles to present two mutually exclusive options (for example, on or off), where choosing an option results in an immediate action.</span></span>
  
#### <a name="example-toggle-in-a-task-pane"></a><span data-ttu-id="3009d-105">Exemple : bouton bascule dans un volet Office</span><span class="sxs-lookup"><span data-stu-id="3009d-105">Example: Toggle in a task pane</span></span>

![Image illustrant le composant de bouton bascule](../images/overview-with-app-toggle.png)

## <a name="best-practices"></a><span data-ttu-id="3009d-107">Meilleures pratiques</span><span class="sxs-lookup"><span data-stu-id="3009d-107">Best practices</span></span>

|<span data-ttu-id="3009d-108">**À faire**</span><span class="sxs-lookup"><span data-stu-id="3009d-108">**Do**</span></span>|<span data-ttu-id="3009d-109">**À ne pas faire**</span><span class="sxs-lookup"><span data-stu-id="3009d-109">**Don't**</span></span>|
|:------------|:--------------|
|<span data-ttu-id="3009d-110">Utiliser les boutons bascule pour les paramètres binaires lorsque les modifications sont immédiatement appliquées.</span><span class="sxs-lookup"><span data-stu-id="3009d-110">Use toggles for binary settings when changes are immediately applied.</span></span><br/><br/>![Exemple de bouton bascule À faire](../images/toggle-do.png)<br/>|<span data-ttu-id="3009d-112">Ne pas utiliser de boutons bascule si les utilisateurs doivent effectuer une étape supplémentaire avant que les modifications prennent effet.</span><span class="sxs-lookup"><span data-stu-id="3009d-112">Don’t use toggles if users must perform an extra step before changes take effect.</span></span><br/><br/>![Exemple de bouton bascule À ne pas faire](../images/toggle-dont.png)<br/>|
|<span data-ttu-id="3009d-p102">Remplacer les étiquettes **On** et **Off** uniquement s’il existe des étiquettes plus spécifiques à utiliser pour un paramètre. Utiliser des étiquettes courtes (3 à 4 caractères) qui représentent des opposés binaires.</span><span class="sxs-lookup"><span data-stu-id="3009d-p102">Only replace the **On** and **Off** labels if there are more specific labels to use for a setting. Use short (3-4 character) labels that represent binary opposites.</span></span>| |

## <a name="variants"></a><span data-ttu-id="3009d-116">Variantes</span><span class="sxs-lookup"><span data-stu-id="3009d-116">Variants</span></span>

|<span data-ttu-id="3009d-117">**Variation**</span><span class="sxs-lookup"><span data-stu-id="3009d-117">**Variation**</span></span>|<span data-ttu-id="3009d-118">**Description**</span><span class="sxs-lookup"><span data-stu-id="3009d-118">**Description**</span></span>|<span data-ttu-id="3009d-119">**Exemple**</span><span class="sxs-lookup"><span data-stu-id="3009d-119">**Example**</span></span>|
|:------------|:--------------|:----------|
|<span data-ttu-id="3009d-120">**Enabled and checked (Activé et sélectionné)**</span><span class="sxs-lookup"><span data-stu-id="3009d-120">**Enabled and checked**</span></span>|<span data-ttu-id="3009d-121">À utiliser lorsque l’état basculé est actif.</span><span class="sxs-lookup"><span data-stu-id="3009d-121">Use when the toggled state is active.</span></span>|![Image Enabled and checked (Activé et sélectionné)](../images/toggle-enabled-on.png)<br/>|
|<span data-ttu-id="3009d-123">**Enabled and unchecked (Activé et désélectionné)**</span><span class="sxs-lookup"><span data-stu-id="3009d-123">**Enabled and unchecked**</span></span>|<span data-ttu-id="3009d-124">À utiliser lorsque l’état basculé est inactif.</span><span class="sxs-lookup"><span data-stu-id="3009d-124">Use when the toggled state is inactive.</span></span>|![Image Enabled and unchecked (Activé et désélectionné)](../images/toggle-enabled-off.png)<br/>|
|<span data-ttu-id="3009d-126">**Disabled and checked (Désactivé et sélectionné)**</span><span class="sxs-lookup"><span data-stu-id="3009d-126">**Disabled and checked**</span></span>|<span data-ttu-id="3009d-127">À utiliser lorsque l’état actif ne peut pas être modifié.</span><span class="sxs-lookup"><span data-stu-id="3009d-127">Use when the active state cannot be changed.</span></span>|![Image Disabled and checked (Désactivé et sélectionné)](../images/toggle-disabled-on.png)<br/>|
|<span data-ttu-id="3009d-129">**Disabled and unchecked (Désactivé et désélectionné)**</span><span class="sxs-lookup"><span data-stu-id="3009d-129">**Disabled and unchecked**</span></span>|<span data-ttu-id="3009d-130">À utiliser lorsque l’état inactif ne peut pas être modifié.</span><span class="sxs-lookup"><span data-stu-id="3009d-130">Use when the inactive state cannot be changed.</span></span>|![Image Disabled and unchecked (Désactivé et désélectionné)](../images/toggle-disabled-off.png)<br/>|

## <a name="implementation"></a><span data-ttu-id="3009d-132">Implémentation</span><span class="sxs-lookup"><span data-stu-id="3009d-132">Implementation</span></span>

<span data-ttu-id="3009d-133">Pour plus d’informations, reportez-vous à [Bouton bascule](https://dev.office.com/fabric#/components/toggle) et [Démarrer avec un exemple de code Fabric React](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).</span><span class="sxs-lookup"><span data-stu-id="3009d-133">For details, see [Toggle](https://dev.office.com/fabric#/components/toggle) and [Getting started with Fabric React code sample](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).</span></span>

## <a name="see-also"></a><span data-ttu-id="3009d-134">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="3009d-134">See also</span></span>

- [<span data-ttu-id="3009d-135">Modèles de conception UX</span><span class="sxs-lookup"><span data-stu-id="3009d-135">UX Design Patterns</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [<span data-ttu-id="3009d-136">Office UI Fabric dans des compléments Office</span><span class="sxs-lookup"><span data-stu-id="3009d-136">Office UI Fabric in Office Add-ins</span></span>](office-ui-fabric.md)
