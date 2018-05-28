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
# <a name="toggle-component-in-office-ui-fabric"></a>Composant de bouton bascule dans Office UI Fabric

Les boutons bascules sont des commutateurs physiques qui activent ou d?sactivent des ?l?ments. Utilisez les boutons bascules pour pr?senter deux options qui s?excluent mutuellement (par exemple, on et off), lorsque le choix d?une option provoque une action imm?diate.
  
#### <a name="example-toggle-in-a-task-pane"></a>Exemple : bouton bascule dans un volet Office

![Image illustrant le composant de bouton bascule](../images/overview-with-app-toggle.png)

## <a name="best-practices"></a>Meilleures pratiques

|**? faire**|**? ne pas faire**|
|:------------|:--------------|
|Utiliser les boutons bascule pour les param?tres binaires lorsque les modifications sont imm?diatement appliqu?es.<br/><br/>![Exemple de bouton bascule ? faire](../images/toggle-do.png)<br/>|Ne pas utiliser de boutons bascule si les utilisateurs doivent effectuer une ?tape suppl?mentaire avant que les modifications prennent effet.<br/><br/>![Exemple de bouton bascule ? ne pas faire](../images/toggle-dont.png)<br/>|
|Remplacer les ?tiquettes **On** et **Off** uniquement s?il existe des ?tiquettes plus sp?cifiques ? utiliser pour un param?tre. Utiliser des ?tiquettes courtes (3 ? 4 caract?res) qui repr?sentent des oppos?s binaires.| |

## <a name="variants"></a>Variantes

|**Variation**|**Description**|**Exemple**|
|:------------|:--------------|:----------|
|**Enabled and checked (Activ? et s?lectionn?)**|? utiliser lorsque l??tat bascul? est actif.|![Image Enabled and checked (Activ? et s?lectionn?)](../images/toggle-enabled-on.png)<br/>|
|**Enabled and unchecked (Activ? et d?s?lectionn?)**|? utiliser lorsque l??tat bascul? est inactif.|![Image Enabled and unchecked (Activ? et d?s?lectionn?)](../images/toggle-enabled-off.png)<br/>|
|**Disabled and checked (D?sactiv? et s?lectionn?)**|? utiliser lorsque l??tat actif ne peut pas ?tre modifi?.|![Image Disabled and checked (D?sactiv? et s?lectionn?)](../images/toggle-disabled-on.png)<br/>|
|**Disabled and unchecked (D?sactiv? et d?s?lectionn?)**|? utiliser lorsque l??tat inactif ne peut pas ?tre modifi?.|![Image Disabled and unchecked (D?sactiv? et d?s?lectionn?)](../images/toggle-disabled-off.png)<br/>|

## <a name="implementation"></a>Impl?mentation

Pour plus d?informations, reportez-vous ? [Bouton bascule](https://dev.office.com/fabric#/components/toggle) et [D?marrer avec un exemple de code Fabric React](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).

## <a name="see-also"></a>Voir aussi

- [Mod?les de conception UX](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [Office UI Fabric dans des compl?ments Office](office-ui-fabric.md)
