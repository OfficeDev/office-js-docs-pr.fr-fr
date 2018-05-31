---
title: Composant TextField dans Office UI Fabric
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 7c579bc12ed0cf1f4d4af52306c6556f7f79f427
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437289"
---
# <a name="textfield-component-in-office-ui-fabric"></a>Composant TextField dans Office UI Fabric

Les champs de texte permettent aux utilisateurs de saisir du texte. Ils sont généralement utilisés pour capturer une seule ligne de texte mais peuvent être configurés pour capturer plusieurs lignes de texte. Le texte s’affiche à l’écran dans un format simple et uniforme.
  
#### <a name="example-textfield-in-a-task-pane"></a>Exemple : TextField dans un volet Office

![Image illustrant le composant TextField](../images/overview-with-app-text-field.png)

## <a name="best-practices"></a>Meilleures pratiques

|**À faire**|**À ne pas faire**|
|:------------|:--------------|
|Utiliser des champs de texte pour accepter la saisie de données sur un formulaire ou une page.|Ne pas utiliser de champs de texte pour rendre une copie de base dans un élément de corps d’une page.|
|Étiqueter les champs de texte avec des noms utiles.|Ne pas utiliser de champs de texte pour saisir une date ou une heure. Utiliser plutôt un sélecteur de date et heure.|
|Utiliser un texte de l’espace réservé concis pour spécifier le contenu qui doit être saisi.|Ne pas utiliser de champs de texte si des options d’entrée valides peuvent être prédéfinies. Utiliser plutôt une liste déroulante.|
|Fournir tous les états appropriés pour les champs de texte (statique, pointage, focus, engagé, non disponible, erreur).||
|Marquer clairement les champs obligatoires et facultatifs.||
|Si possible, mettre en forme les champs de texte en fonction du format de données attendu. Par exemple, lors de la capture d’un numéro de téléphone à 10 chiffres, utiliser trois champs distincts pour stocker les différentes parties du numéro de téléphone.||

## <a name="variants"></a>Variantes

|**Variation**|**Description**|**Exemple**|
|:------------|:--------------|:----------|
|**Default TextField (champ de texte par défaut)**|À utiliser comme champ de texte par défaut.|![Image Default TextField (champ de texte par défaut)](../images/textfield-default.png)<br/>|
|**Disabled TextField (champ de texte désactivé)**|À utiliser lorsque le champ de texte est désactivé.|![Image Disabled TextField (champ de texte désactivé)](../images/textfield-disabled.png)<br/>|
|**Required TextField (champ de texte obligatoire)**|À utiliser lorsque le champ de texte est obligatoire.|![Image Required TextField (champ de texte obligatoire)](../images/textfield-required.png)<br/>|
|**TextField with a placeholder (champ de texte avec un espace réservé)**|À utiliser lorsqu’un texte de l’espace réservé est nécessaire.|![Image TextField with a placeholder (champ de texte avec un espace réservé)](../images/textfield-placeholder.png)<br/>|
|**TextField with multiple lines (champ de texte avec plusieurs lignes)**|À utiliser lorsque plusieurs lignes de texte sont nécessaires.|![Image TextField with a placeholder (champ de texte avec un espace réservé)](../images/textfield-multi.png)<br/>|

## <a name="implementation"></a>Implémentation

Pour plus d’informations, reportez-vous à [TextField](https://dev.office.com/fabric#/components/textfield) et [Démarrer avec un exemple de code Fabric React](https://github.com/OfficeDev/Word-Add-in-GettingStartedFabricReact).

## <a name="see-also"></a>Voir aussi

- [Modèles de conception UX](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
- [Office UI Fabric dans des compléments Office](office-ui-fabric.md)
