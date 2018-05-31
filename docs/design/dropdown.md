---
title: Composant de liste déroulante dans Office UI Fabric
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: af780d5c88c95eb742f82b59e522e4c34a276d31
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437387"
---
# <a name="dropdown-component-in-office-ui-fabric"></a><span data-ttu-id="3b27a-102">Composant de liste déroulante dans Office UI Fabric</span><span class="sxs-lookup"><span data-stu-id="3b27a-102">DropDown component in Office UI Fabric</span></span>

<span data-ttu-id="3b27a-p101">Une liste déroulante est une liste d’options qui s’affiche en cliquant sur un bouton de liste déroulante. Utilisez des listes déroulantes ou des menus déroulants pour simplifier la conception de l’interface utilisateur, et lorsque les utilisateurs doivent effectuer un choix dans l’interface utilisateur. Lorsque la liste est réduite, l’élément sélectionné est visible. Pour modifier l’élément sélectionné, les utilisateurs ouvrent la liste et sélectionnent une nouvelle valeur.</span><span class="sxs-lookup"><span data-stu-id="3b27a-p101">A drop-down is a list of options that is shown by clicking a drop-down button. Use a drop-down list or menu to simplify the UI design, and when users should make a choice within the UI. When the list collapses, the selected item is visible. To change the selected item, users open the list, and select a new value.</span></span>
  
#### <a name="example-drop-down-in-a-task-pane"></a><span data-ttu-id="3b27a-107">Exemple : Liste déroulante dans un volet de tâches</span><span class="sxs-lookup"><span data-stu-id="3b27a-107">Example: Drop-down in a task pane</span></span>

![Image illustrant la liste déroulante](../images/overview-with-app-dropdown.png)

## <a name="best-practices"></a><span data-ttu-id="3b27a-109">Meilleures pratiques</span><span class="sxs-lookup"><span data-stu-id="3b27a-109">Best practices</span></span>

|<span data-ttu-id="3b27a-110">**À faire**</span><span class="sxs-lookup"><span data-stu-id="3b27a-110">**Do**</span></span>|<span data-ttu-id="3b27a-111">**À ne pas faire**</span><span class="sxs-lookup"><span data-stu-id="3b27a-111">**Don't**</span></span>|
|:------------|:--------------|
|<span data-ttu-id="3b27a-p102">Utiliser une liste déroulante lorsque l’option sélectionnée par défaut est plus susceptible d’être sélectionnée que d’autres options. En revanche, ChoiceGroup ou les cases d’option affichent tous les choix, ce qui donne une importance identique à toutes les options.</span><span class="sxs-lookup"><span data-stu-id="3b27a-p102">Use a drop-down when the default selected option is more likely to be selected than other options. By contrast, ChoiceGroup or radio buttons show all choices, thereby putting equal emphasis on all options.</span></span>|<span data-ttu-id="3b27a-114">Ne pas utiliser de listes déroulantes lorsque toutes les options sont pareillement susceptibles d’être sélectionnées.</span><span class="sxs-lookup"><span data-stu-id="3b27a-114">Don't use a drop-down when all options are equally likely to be selected.</span></span>|
|<span data-ttu-id="3b27a-p103">Utiliser une liste déroulante lorsqu’il y a plusieurs choix qui peuvent être réduits en un seul champ. Utiliser également les listes déroulantes pour les longues listes d’éléments ou lorsque l’espace à l’écran est limité.</span><span class="sxs-lookup"><span data-stu-id="3b27a-p103">Use a drop-down when there are multiple choices that can be collapsed into one field. Also, use a drop-down for long lists of items, or when screen space is constrained.</span></span>|<span data-ttu-id="3b27a-p104">Ne pas utiliser de listes déroulantes s’il y a moins de deux choix. À la place, utiliser une case à cocher.</span><span class="sxs-lookup"><span data-stu-id="3b27a-p104">Don’t use a drop-down if there are fewer than two choices. Instead, use a check box.</span></span>|
|<span data-ttu-id="3b27a-119">Utiliser des instructions ou des mots raccourcis dans les listes déroulantes.</span><span class="sxs-lookup"><span data-stu-id="3b27a-119">Use shortened statements or words in a drop-down.</span></span>| |

## <a name="variants"></a><span data-ttu-id="3b27a-120">Variantes</span><span class="sxs-lookup"><span data-stu-id="3b27a-120">Variants</span></span>

|<span data-ttu-id="3b27a-121">**Variation**</span><span class="sxs-lookup"><span data-stu-id="3b27a-121">**Variation**</span></span>|<span data-ttu-id="3b27a-122">**Description**</span><span class="sxs-lookup"><span data-stu-id="3b27a-122">**Description**</span></span>|<span data-ttu-id="3b27a-123">**Exemple**</span><span class="sxs-lookup"><span data-stu-id="3b27a-123">**Example**</span></span>|
|:------------|:--------------|:----------|
|<span data-ttu-id="3b27a-124">**Liste déroulante non contrôlée de base**</span><span class="sxs-lookup"><span data-stu-id="3b27a-124">**Basic uncontrolled drop-down**</span></span>|<span data-ttu-id="3b27a-125">À utiliser lorsque de nombreuses options peuvent être sélectionnées.</span><span class="sxs-lookup"><span data-stu-id="3b27a-125">Use when many options are available for selection.</span></span>|![Image de la liste déroulante non contrôlée de base](../images/dropdown-uncontrolled.png)<br/>|
|<span data-ttu-id="3b27a-127">**Liste déroulante non contrôlée désactivée avec defaultSelectedKey**</span><span class="sxs-lookup"><span data-stu-id="3b27a-127">**Disabled uncontrolled drop-down with defaultSelectedKey**</span></span>|<span data-ttu-id="3b27a-128">État désactivé de la liste déroulante.</span><span class="sxs-lookup"><span data-stu-id="3b27a-128">Disabled state of the drop-down.</span></span>|![Image de la liste déroulante non contrôlée désactivée avec defaultSelectedKey](../images/dropdown-disabled.png)<br/>|
|<span data-ttu-id="3b27a-130">**Liste déroulante contrôlée**</span><span class="sxs-lookup"><span data-stu-id="3b27a-130">**Controlled drop-down**</span></span>|<span data-ttu-id="3b27a-131">À utiliser lorsque l’élément sélectionné par défaut est influencé par un autre emplacement dans votre interface utilisateur et que l’élément sélectionné dans la liste déroulante doit être conservé.</span><span class="sxs-lookup"><span data-stu-id="3b27a-131">Use when the default selected item is influenced by another location in your UI, and the selected item in the drop-down must be maintained.</span></span>|![Image de la liste déroulante contrôlée](../images/dropdown-controlled.png)<br/>|

## <a name="implementation"></a><span data-ttu-id="3b27a-133">Implémentation</span><span class="sxs-lookup"><span data-stu-id="3b27a-133">Implementation</span></span>

<span data-ttu-id="3b27a-134">Pour plus d’informations, reportez-vous à [Liste déroulante](https://dev.office.com/fabric#/components/dropdown) et [Démarrer avec un exemple de code Fabric React](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).</span><span class="sxs-lookup"><span data-stu-id="3b27a-134">For details, see [Dropdown](https://dev.office.com/fabric#/components/dropdown) and [Getting started with Fabric React code sample](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).</span></span>

## <a name="see-also"></a><span data-ttu-id="3b27a-135">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="3b27a-135">See also</span></span>

- [<span data-ttu-id="3b27a-136">Modèles de conception UX</span><span class="sxs-lookup"><span data-stu-id="3b27a-136">UX Design Patterns</span></span>](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [<span data-ttu-id="3b27a-137">Office UI Fabric dans des compléments Office</span><span class="sxs-lookup"><span data-stu-id="3b27a-137">Office UI Fabric in Office Add-ins</span></span>](office-ui-fabric.md)
