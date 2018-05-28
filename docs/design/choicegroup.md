---
title: Composant ChoiceGroup dans Office UI Fabric
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 78da2fae781039663bfe2bac159bfbe50192c023
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="choicegroup-component-in-office-ui-fabric"></a><span data-ttu-id="8d2e7-102">Composant ChoiceGroup dans Office UI Fabric</span><span class="sxs-lookup"><span data-stu-id="8d2e7-102">ChoiceGroup component in Office UI Fabric</span></span>

<span data-ttu-id="8d2e7-p101">Le composant ChoiceGroup, ?galement appel? bouton radio, pr?sente aux utilisateurs deux options ou plus qui s?excluent mutuellement. Les utilisateurs ne peuvent s?lectionner qu?un seul bouton ChoiceGroup dans un groupe. Chaque option est repr?sent?e par un bouton ChoiceGroup.</span><span class="sxs-lookup"><span data-stu-id="8d2e7-p101">The ChoiceGroup component, also known as a radio button, presents users with two or more mutually exclusive options. Users can select only one ChoiceGroup button in a group. Each option is represented by one ChoiceGroup button.</span></span> 
  
#### <a name="example-choicegroup-in-a-task-pane"></a><span data-ttu-id="8d2e7-106">Exemple : ChoiceGroup dans un volet des t?ches</span><span class="sxs-lookup"><span data-stu-id="8d2e7-106">Example: ChoiceGroup in a task pane</span></span>

 ![Image illustrant un ChoiceGroup](../images/overview-with-app-choicegroup.png)

## <a name="best-practices"></a><span data-ttu-id="8d2e7-108">Meilleures pratiques</span><span class="sxs-lookup"><span data-stu-id="8d2e7-108">Best practices</span></span>

|<span data-ttu-id="8d2e7-109">**? faire**</span><span class="sxs-lookup"><span data-stu-id="8d2e7-109">**Do**</span></span>|<span data-ttu-id="8d2e7-110">**? ne pas faire**</span><span class="sxs-lookup"><span data-stu-id="8d2e7-110">**Don't**</span></span>|
|:------------|:--------------|
|<span data-ttu-id="8d2e7-111">Conserver les options ChoiceGroup au m?me niveau.</span><span class="sxs-lookup"><span data-stu-id="8d2e7-111">Keep ChoiceGroup options at the same level.</span></span><br/><br/>![Exemple ChoiceGroup ? faire](../images/choice-do.png)<br/>|<span data-ttu-id="8d2e7-113">Ne pas utiliser de ChoiceGroups ou de cases ? cocher imbriqu?s.</span><span class="sxs-lookup"><span data-stu-id="8d2e7-113">Don't use nested ChoiceGroups or check boxes.</span></span><br/><br/>![Exemple ChoiceGroup ? ne pas faire](../images/choice-dont.png)<br/>|
|<span data-ttu-id="8d2e7-115">Utiliser des ChoiceGroups avec 2 ? 7 options, en v?rifiant qu?il y a suffisamment d?espace ? l??cran pour afficher toutes les options.</span><span class="sxs-lookup"><span data-stu-id="8d2e7-115">Use ChoiceGroups with 2-7 options, ensuring there is enough screen space to show all options.</span></span> <span data-ttu-id="8d2e7-116">Dans le cas contraire, utiliser une case ? cocher ou une liste d?roulante.</span><span class="sxs-lookup"><span data-stu-id="8d2e7-116">Otherwise, use a check box or drop-down list.</span></span>|<span data-ttu-id="8d2e7-p103">Ne pas utiliser lorsque les options sont des nombres avec un intervalle fixe, par exemple, 10, 20, 30 et ainsi de suite. ? la place, utiliser un composant de curseur.</span><span class="sxs-lookup"><span data-stu-id="8d2e7-p103">Don't use when the options are numbers with a fixed step, for example 10, 20, 30, and so on. Instead, use a slider component.</span></span>|
|<span data-ttu-id="8d2e7-119">Si les utilisateurs ne choisissent aucune option, envisager d?inclure une option comme **Aucune** ou **Non concern?**.</span><span class="sxs-lookup"><span data-stu-id="8d2e7-119">If users may not choose any of the options, consider including an option such as **None** or **Does not apply**.</span></span>|<span data-ttu-id="8d2e7-120">Ne pas utiliser de boutons ChoiceGroup pour un choix binaire unique.</span><span class="sxs-lookup"><span data-stu-id="8d2e7-120">Don?t use two ChoiceGroup buttons for a single binary choice.</span></span>|
|<span data-ttu-id="8d2e7-p104">Si possible, aligner les boutons ChoiceGroup verticalement et non horizontalement. L?alignement horizontal est plus difficile ? lire et ? localiser.</span><span class="sxs-lookup"><span data-stu-id="8d2e7-p104">If possible, align ChoiceGroup buttons vertically instead of horizontally. Horizontal alignment is harder to read and localize.</span></span>||
|<span data-ttu-id="8d2e7-123">Lister les options dans un ordre logique. Par exemple, commencer par les options les plus susceptibles d??tre activ?es, les plus simples ou les moins risqu?es.</span><span class="sxs-lookup"><span data-stu-id="8d2e7-123">List options in logical order, for example, the most likely option to be selected to the least, the simplest operation to the most complex, or the least risk to the highest risk.</span></span> |<span data-ttu-id="8d2e7-124">Ne pas ranger les options par ordre alphab?tique, car ce classement d?pend de la langue.</span><span class="sxs-lookup"><span data-stu-id="8d2e7-124">Don't use alphabetical ordering because it is language dependent.</span></span>|

## <a name="variants"></a><span data-ttu-id="8d2e7-125">Variantes</span><span class="sxs-lookup"><span data-stu-id="8d2e7-125">Variants</span></span>

|<span data-ttu-id="8d2e7-126">**Variation**</span><span class="sxs-lookup"><span data-stu-id="8d2e7-126">**Variation**</span></span>|<span data-ttu-id="8d2e7-127">**Description**</span><span class="sxs-lookup"><span data-stu-id="8d2e7-127">**Description**</span></span>|<span data-ttu-id="8d2e7-128">**Exemple**</span><span class="sxs-lookup"><span data-stu-id="8d2e7-128">**Example**</span></span>|
|:------------|:--------------|:----------|
|<span data-ttu-id="8d2e7-129">**ChoiceGroups**</span><span class="sxs-lookup"><span data-stu-id="8d2e7-129">**ChoiceGroups**</span></span>|<span data-ttu-id="8d2e7-130">? utiliser lorsque les images ne sont pas n?cessaires pour effectuer une s?lection.</span><span class="sxs-lookup"><span data-stu-id="8d2e7-130">Use when imagery is not necessary for making a selection.</span></span>|![Image de variante ChoiceGroup](../images/radio.png)<br/>|
|<span data-ttu-id="8d2e7-132">**ChoiceGroups utilisant des images**</span><span class="sxs-lookup"><span data-stu-id="8d2e7-132">**ChoiceGroups using images**</span></span>|<span data-ttu-id="8d2e7-133">? utiliser lorsque les images sont n?cessaires pour effectuer une s?lection.</span><span class="sxs-lookup"><span data-stu-id="8d2e7-133">Use when imagery is necessary for making a selection.</span></span>|![Variante ChoiceGroup avec image](../images/radio-image.png)<br/>|

## <a name="implementation"></a><span data-ttu-id="8d2e7-135">Impl?mentation</span><span class="sxs-lookup"><span data-stu-id="8d2e7-135">Implementation</span></span>

<span data-ttu-id="8d2e7-136">Pour plus d?informations, reportez-vous ? [ChoiceGroup](https://dev.office.com/fabric#/components/choicegroup) et [D?marrer avec un exemple de code Fabric React](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).</span><span class="sxs-lookup"><span data-stu-id="8d2e7-136">For details, see [ChoiceGroup](https://dev.office.com/fabric#/components/choicegroup) and [Getting started with Fabric React code sample](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).</span></span>

## <a name="see-also"></a><span data-ttu-id="8d2e7-137">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="8d2e7-137">See also</span></span>

- [<span data-ttu-id="8d2e7-138">Mod?les de conception UX</span><span class="sxs-lookup"><span data-stu-id="8d2e7-138">UX Design Patterns</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [<span data-ttu-id="8d2e7-139">Office UI Fabric dans des compl?ments Office</span><span class="sxs-lookup"><span data-stu-id="8d2e7-139">Office UI Fabric in Office Add-ins</span></span>](office-ui-fabric.md)
