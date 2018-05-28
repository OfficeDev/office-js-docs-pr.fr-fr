---
title: Composant TextField dans Office UI Fabric
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 7c579bc12ed0cf1f4d4af52306c6556f7f79f427
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="textfield-component-in-office-ui-fabric"></a>Composant TextField dans Office UI Fabric

Les champs de texte permettent aux utilisateurs de saisir du texte. Ils sont g?n?ralement utilis?s pour capturer une seule ligne de texte mais peuvent ?tre configur?s pour capturer plusieurs lignes de texte. Le texte s?affiche ? l??cran dans un format simple et uniforme.
  
#### <a name="example-textfield-in-a-task-pane"></a>Exemple : TextField dans un volet Office

![Image illustrant le composant TextField](../images/overview-with-app-text-field.png)

## <a name="best-practices"></a>Meilleures pratiques

|**? faire**|**? ne pas faire**|
|:------------|:--------------|
|Utiliser des champs de texte pour accepter la saisie de donn?es sur un formulaire ou une page.|Ne pas utiliser de champs de texte pour rendre une copie de base dans un ?l?ment de corps d?une page.|
|?tiqueter les champs de texte avec des noms utiles.|Ne pas utiliser de champs de texte pour saisir une date ou une heure. Utiliser plut?t un s?lecteur de date et heure.|
|Utiliser un texte de l?espace r?serv? concis pour sp?cifier le contenu qui doit ?tre saisi.|Ne pas utiliser de champs de texte si des options d?entr?e valides peuvent ?tre pr?d?finies. Utiliser plut?t une liste d?roulante.|
|Fournir tous les ?tats appropri?s pour les champs de texte (statique, pointage, focus, engag?, non disponible, erreur).||
|Marquer clairement les champs obligatoires et facultatifs.||
|Si possible, mettre en forme les champs de texte en fonction du format de donn?es attendu. Par exemple, lors de la capture d?un num?ro de t?l?phone ? 10 chiffres, utiliser trois champs distincts pour stocker les diff?rentes parties du num?ro de t?l?phone.||

## <a name="variants"></a>Variantes

|**Variation**|**Description**|**Exemple**|
|:------------|:--------------|:----------|
|**Default TextField (champ de texte par d?faut)**|? utiliser comme champ de texte par d?faut.|![Image Default TextField (champ de texte par d?faut)](../images/textfield-default.png)<br/>|
|**Disabled TextField (champ de texte d?sactiv?)**|? utiliser lorsque le champ de texte est d?sactiv?.|![Image Disabled TextField (champ de texte d?sactiv?)](../images/textfield-disabled.png)<br/>|
|**Required TextField (champ de texte obligatoire)**|? utiliser lorsque le champ de texte est obligatoire.|![Image Required TextField (champ de texte obligatoire)](../images/textfield-required.png)<br/>|
|**TextField with a placeholder (champ de texte avec un espace r?serv?)**|? utiliser lorsqu?un texte de l?espace r?serv? est n?cessaire.|![Image TextField with a placeholder (champ de texte avec un espace r?serv?)](../images/textfield-placeholder.png)<br/>|
|**TextField with multiple lines (champ de texte avec plusieurs lignes)**|? utiliser lorsque plusieurs lignes de texte sont n?cessaires.|![Image TextField with a placeholder (champ de texte avec un espace r?serv?)](../images/textfield-multi.png)<br/>|

## <a name="implementation"></a>Impl?mentation

Pour plus d?informations, reportez-vous ? [TextField](https://dev.office.com/fabric#/components/textfield) et [D?marrer avec un exemple de code Fabric React](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).

## <a name="see-also"></a>Voir aussi

- [Mod?les de conception UX](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [Office UI Fabric dans des compl?ments Office](office-ui-fabric.md)
