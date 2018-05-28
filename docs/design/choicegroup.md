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
# <a name="choicegroup-component-in-office-ui-fabric"></a>Composant ChoiceGroup dans Office UI Fabric

Le composant ChoiceGroup, ?galement appel? bouton radio, pr?sente aux utilisateurs deux options ou plus qui s?excluent mutuellement. Les utilisateurs ne peuvent s?lectionner qu?un seul bouton ChoiceGroup dans un groupe. Chaque option est repr?sent?e par un bouton ChoiceGroup. 
  
#### <a name="example-choicegroup-in-a-task-pane"></a>Exemple : ChoiceGroup dans un volet des t?ches

 ![Image illustrant un ChoiceGroup](../images/overview-with-app-choicegroup.png)

## <a name="best-practices"></a>Meilleures pratiques

|**? faire**|**? ne pas faire**|
|:------------|:--------------|
|Conserver les options ChoiceGroup au m?me niveau.<br/><br/>![Exemple ChoiceGroup ? faire](../images/choice-do.png)<br/>|Ne pas utiliser de ChoiceGroups ou de cases ? cocher imbriqu?s.<br/><br/>![Exemple ChoiceGroup ? ne pas faire](../images/choice-dont.png)<br/>|
|Utiliser des ChoiceGroups avec 2 ? 7 options, en v?rifiant qu?il y a suffisamment d?espace ? l??cran pour afficher toutes les options. Dans le cas contraire, utiliser une case ? cocher ou une liste d?roulante.|Ne pas utiliser lorsque les options sont des nombres avec un intervalle fixe, par exemple, 10, 20, 30 et ainsi de suite. ? la place, utiliser un composant de curseur.|
|Si les utilisateurs ne choisissent aucune option, envisager d?inclure une option comme **Aucune** ou **Non concern?**.|Ne pas utiliser de boutons ChoiceGroup pour un choix binaire unique.|
|Si possible, aligner les boutons ChoiceGroup verticalement et non horizontalement. L?alignement horizontal est plus difficile ? lire et ? localiser.||
|Lister les options dans un ordre logique. Par exemple, commencer par les options les plus susceptibles d??tre activ?es, les plus simples ou les moins risqu?es. |Ne pas ranger les options par ordre alphab?tique, car ce classement d?pend de la langue.|

## <a name="variants"></a>Variantes

|**Variation**|**Description**|**Exemple**|
|:------------|:--------------|:----------|
|**ChoiceGroups**|? utiliser lorsque les images ne sont pas n?cessaires pour effectuer une s?lection.|![Image de variante ChoiceGroup](../images/radio.png)<br/>|
|**ChoiceGroups utilisant des images**|? utiliser lorsque les images sont n?cessaires pour effectuer une s?lection.|![Variante ChoiceGroup avec image](../images/radio-image.png)<br/>|

## <a name="implementation"></a>Impl?mentation

Pour plus d?informations, reportez-vous ? [ChoiceGroup](https://dev.office.com/fabric#/components/choicegroup) et [D?marrer avec un exemple de code Fabric React](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).

## <a name="see-also"></a>Voir aussi

- [Mod?les de conception UX](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [Office UI Fabric dans des compl?ments Office](office-ui-fabric.md)
