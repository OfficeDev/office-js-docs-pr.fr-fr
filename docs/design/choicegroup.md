---
title: Composant ChoiceGroup dans Office UI Fabric
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 78da2fae781039663bfe2bac159bfbe50192c023
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437212"
---
# <a name="choicegroup-component-in-office-ui-fabric"></a><span data-ttu-id="f2f40-102">Composant ChoiceGroup dans Office UI Fabric</span><span class="sxs-lookup"><span data-stu-id="f2f40-102">ChoiceGroup component in Office UI Fabric</span></span>

<span data-ttu-id="f2f40-p101">Le composant ChoiceGroup, également appelé bouton radio, présente aux utilisateurs deux options ou plus qui s’excluent mutuellement. Les utilisateurs ne peuvent sélectionner qu’un seul bouton ChoiceGroup dans un groupe. Chaque option est représentée par un bouton ChoiceGroup.</span><span class="sxs-lookup"><span data-stu-id="f2f40-p101">The ChoiceGroup component, also known as a radio button, presents users with two or more mutually exclusive options. Users can select only one ChoiceGroup button in a group. Each option is represented by one ChoiceGroup button.</span></span> 
  
#### <a name="example-choicegroup-in-a-task-pane"></a><span data-ttu-id="f2f40-106">Exemple : ChoiceGroup dans un volet des tâches</span><span class="sxs-lookup"><span data-stu-id="f2f40-106">Example: ChoiceGroup in a task pane</span></span>

 ![Image illustrant un ChoiceGroup](../images/overview-with-app-choicegroup.png)

## <a name="best-practices"></a><span data-ttu-id="f2f40-108">Meilleures pratiques</span><span class="sxs-lookup"><span data-stu-id="f2f40-108">Best practices</span></span>

|<span data-ttu-id="f2f40-109">**À faire**</span><span class="sxs-lookup"><span data-stu-id="f2f40-109">**Do**</span></span>|<span data-ttu-id="f2f40-110">**À ne pas faire**</span><span class="sxs-lookup"><span data-stu-id="f2f40-110">**Don't**</span></span>|
|:------------|:--------------|
|<span data-ttu-id="f2f40-111">Conserver les options ChoiceGroup au même niveau.</span><span class="sxs-lookup"><span data-stu-id="f2f40-111">Keep ChoiceGroup options at the same level.</span></span><br/><br/>![Exemple ChoiceGroup À faire](../images/choice-do.png)<br/>|<span data-ttu-id="f2f40-113">Ne pas utiliser de ChoiceGroups ou de cases à cocher imbriqués.</span><span class="sxs-lookup"><span data-stu-id="f2f40-113">Don't use nested ChoiceGroups or check boxes.</span></span><br/><br/>![Exemple ChoiceGroup À ne pas faire](../images/choice-dont.png)<br/>|
|<span data-ttu-id="f2f40-115">Utiliser des ChoiceGroups avec 2 à 7 options, en vérifiant qu’il y a suffisamment d’espace à l’écran pour afficher toutes les options.</span><span class="sxs-lookup"><span data-stu-id="f2f40-115">Use ChoiceGroups with 2-7 options, ensuring there is enough screen space to show all options.</span></span> <span data-ttu-id="f2f40-116">Dans le cas contraire, utiliser une case à cocher ou une liste déroulante.</span><span class="sxs-lookup"><span data-stu-id="f2f40-116">Otherwise, use a check box or drop-down list.</span></span>|<span data-ttu-id="f2f40-p103">Ne pas utiliser lorsque les options sont des nombres avec un intervalle fixe, par exemple, 10, 20, 30 et ainsi de suite. À la place, utiliser un composant de curseur.</span><span class="sxs-lookup"><span data-stu-id="f2f40-p103">Don't use when the options are numbers with a fixed step, for example 10, 20, 30, and so on. Instead, use a slider component.</span></span>|
|<span data-ttu-id="f2f40-119">Si les utilisateurs ne choisissent aucune option, envisager d’inclure une option comme **Aucune** ou **Non concerné**.</span><span class="sxs-lookup"><span data-stu-id="f2f40-119">If users may not choose any of the options, consider including an option such as **None** or **Does not apply**.</span></span>|<span data-ttu-id="f2f40-120">Ne pas utiliser de boutons ChoiceGroup pour un choix binaire unique.</span><span class="sxs-lookup"><span data-stu-id="f2f40-120">Don’t use two ChoiceGroup buttons for a single binary choice.</span></span>|
|<span data-ttu-id="f2f40-p104">Si possible, aligner les boutons ChoiceGroup verticalement et non horizontalement. L’alignement horizontal est plus difficile à lire et à localiser.</span><span class="sxs-lookup"><span data-stu-id="f2f40-p104">If possible, align ChoiceGroup buttons vertically instead of horizontally. Horizontal alignment is harder to read and localize.</span></span>||
|<span data-ttu-id="f2f40-123">Lister les options dans un ordre logique. Par exemple, commencer par les options les plus susceptibles d’être activées, les plus simples ou les moins risquées.</span><span class="sxs-lookup"><span data-stu-id="f2f40-123">List options in logical order, for example, the most likely option to be selected to the least, the simplest operation to the most complex, or the least risk to the highest risk.</span></span> |<span data-ttu-id="f2f40-124">Ne pas ranger les options par ordre alphabétique, car ce classement dépend de la langue.</span><span class="sxs-lookup"><span data-stu-id="f2f40-124">Don't use alphabetical ordering because it is language dependent.</span></span>|

## <a name="variants"></a><span data-ttu-id="f2f40-125">Variantes</span><span class="sxs-lookup"><span data-stu-id="f2f40-125">Variants</span></span>

|<span data-ttu-id="f2f40-126">**Variation**</span><span class="sxs-lookup"><span data-stu-id="f2f40-126">**Variation**</span></span>|<span data-ttu-id="f2f40-127">**Description**</span><span class="sxs-lookup"><span data-stu-id="f2f40-127">**Description**</span></span>|<span data-ttu-id="f2f40-128">**Exemple**</span><span class="sxs-lookup"><span data-stu-id="f2f40-128">**Example**</span></span>|
|:------------|:--------------|:----------|
|<span data-ttu-id="f2f40-129">**ChoiceGroups**</span><span class="sxs-lookup"><span data-stu-id="f2f40-129">**ChoiceGroups**</span></span>|<span data-ttu-id="f2f40-130">À utiliser lorsque les images ne sont pas nécessaires pour effectuer une sélection.</span><span class="sxs-lookup"><span data-stu-id="f2f40-130">Use when imagery is not necessary for making a selection.</span></span>|![Image de variante ChoiceGroup](../images/radio.png)<br/>|
|<span data-ttu-id="f2f40-132">**ChoiceGroups utilisant des images**</span><span class="sxs-lookup"><span data-stu-id="f2f40-132">**ChoiceGroups using images**</span></span>|<span data-ttu-id="f2f40-133">À utiliser lorsque les images sont nécessaires pour effectuer une sélection.</span><span class="sxs-lookup"><span data-stu-id="f2f40-133">Use when imagery is necessary for making a selection.</span></span>|![Variante ChoiceGroup avec image](../images/radio-image.png)<br/>|

## <a name="implementation"></a><span data-ttu-id="f2f40-135">Implémentation</span><span class="sxs-lookup"><span data-stu-id="f2f40-135">Implementation</span></span>

<span data-ttu-id="f2f40-136">Pour plus d’informations, reportez-vous à [ChoiceGroup](https://dev.office.com/fabric#/components/choicegroup) et [Démarrer avec un exemple de code Fabric React](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).</span><span class="sxs-lookup"><span data-stu-id="f2f40-136">For details, see [ChoiceGroup](https://dev.office.com/fabric#/components/choicegroup) and [Getting started with Fabric React code sample](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).</span></span>

## <a name="see-also"></a><span data-ttu-id="f2f40-137">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="f2f40-137">See also</span></span>

- [<span data-ttu-id="f2f40-138">Modèles de conception UX</span><span class="sxs-lookup"><span data-stu-id="f2f40-138">UX Design Patterns</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [<span data-ttu-id="f2f40-139">Office UI Fabric dans des compléments Office</span><span class="sxs-lookup"><span data-stu-id="f2f40-139">Office UI Fabric in Office Add-ins</span></span>](office-ui-fabric.md)
