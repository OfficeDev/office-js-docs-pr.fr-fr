---
title: Compléments Outlook contextuels
description: Lancer des tâches liées à un message sans laisser le message lui-même pour faciliter et enrichir l'expérience utilisateur.
ms.date: 10/09/2019
localization_priority: Normal
ms.openlocfilehash: 84ea058e031fd2334706145bcdf8ca8e530c2c38
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720805"
---
# <a name="contextual-outlook-add-ins"></a>Compléments Outlook contextuels

Les compléments contextuels sont des compléments Outlook qui s’activent en fonction du texte d’un message ou d’un rendez-vous. Grâce aux compléments contextuels, vous pouvez initier des tâches associées à un message sans avoir à quitter ce dernier. L’expérience utilisateur en est ainsi facilitée et enrichie.

Voici quelques exemples de compléments contextuels :

- Choix d’une adresse à ouvrir dans un plan du lieu.
- Choix d’une chaîne ouvrant un complément de suggestion de réunion.
- Choisir un numéro de téléphone permet de l’ajouter à vos contacts.


> [!NOTE]
> Les compléments contextuels ne sont pas disponibles actuellement dans Outlook pour Android et iOS. Cette fonctionnalité sera disponible ultérieurement.
>
> La prise en charge de cette fonctionnalité a été introduite dans l’ensemble de conditions requises 1.6. Voir [les clients et les plateformes](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.

## <a name="how-to-make-a-contextual-add-in"></a>Création d’un complément contextuel

Le manifeste d’un complément contextuel doit inclure un élément [ExtensionPoint](../reference/manifest/extensionpoint.md) avec une attribut `xsi:type` défini sur `DetectedEntity`. Au sein de l’élément **ExtensionPoint**, le complément spécifie les entités ou l’expression régulière qui peuvent l’activer. Si une entité est spécifiée, il peut s’agir d’une des propriétés de l’objet [Entités](/javascript/api/outlook/office.entities).

Par conséquent, le manifeste du complément doit contenir un type de règle **ItemHasKnownEntity** ou **Itemhasregularexpressionmatch**. L’exemple suivant montre comment spécifier qu’un complément doit s’activer sur les messages comportant une entité détectée telle qu’un numéro de téléphone :

```XML
<ExtensionPoint xsi:type="DetectedEntity">
  <Label resid="contextLabel" />
  <SourceLocation resid="detectedEntityURL" />
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="PhoneNumber" Highlight="all" />
  </Rule>
</ExtensionPoint>
```

Une fois qu’un complément contextuel est associé à un compte, il démarre automatiquement lorsque l’utilisateur clique sur une expression régulière ou une entité mise en surbrillance. Pour plus d’informations sur les expressions régulières pour les compléments Outlook, reportez-vous à l’article [Utiliser des règles d’activation d’expression régulière pour afficher un complément Outlook](use-regular-expressions-to-show-an-outlook-add-in.md).

Il existe plusieurs restrictions sur les compléments contextuels :

- Un complément contextuel ne peut exister que dans des compléments de lecture (pas dans des compléments de composition).
- Vous ne pouvez pas spécifier la couleur de l’entité en surbrillance.
- Si une entité n’est pas en surbrillance, elle ne lancera pas de complément contextuel dans une carte.

Une entité ou une expression régulière qui n’est pas mise en surbrillance ne permettant pas le lancement d’un complément contextuel, les compléments doivent inclure au moins un élément `Rule` avec l’attribut `Highlight` défini sur `all`.

> [!NOTE]
> Les types d’entité `EmailAddress` et `Url` ne prennent pas en charge la mise en surbrillance. Ils ne peuvent donc pas être utilisés pour lancer un complément contextuel. Ils peuvent toutefois être combinés dans un type de règle `RuleCollection` comme un critère d’activation supplémentaire.

## <a name="how-to-launch-a-contextual-add-in"></a>Lancement d’un complément contextuel

Un utilisateur lance un complément contextuel par le biais du texte, une entité connue ou une expression régulière du développeur. En règle générale, un utilisateur identifie un complément contextuel car l’entité est mise en surbrillance. L’exemple suivant montre comment la mise en surbrillance s’affiche dans un message. Ici, l’entité (une adresse) est colorée en bleu et soulignée avec une ligne bleue en pointillés. Un utilisateur lance le complément contextuel en cliquant sur l’entité en surbrillance. 

**Exemple de texte avec l’entité (une adresse) en surbrillance**

![Illustre l’entité en surbrillance dans un courrier](../images/outlook-detected-entity-highlight.png)
    
Lorsque plusieurs entités ou compléments contextuels sont présents dans un message, l’interaction avec l’utilisateur a lieu comme suit :

- S’il existe plusieurs entités, l’utilisateur doit cliquer sur une autre entité pour lancer le complément de celle-ci.
- Si une entité active plusieurs compléments, chaque complément ouvre un nouvel onglet. L’utilisateur bascule entre les onglets pour passer d’un complément à l’autre. Par exemple, un nom et une adresse peuvent déclencher un complément de téléphone et une carte.
- Si une chaîne unique contient plusieurs entités qui activent plusieurs compléments, la chaîne entière est mise en surbrillance et lorsque l’utilisateur clique sur cette chaîne, tous les compléments concernés par la chaîne s’affichent dans des onglets distincts. Par exemple, une chaîne qui décrit une proposition de réunion dans un restaurant peut activer le complément de suggestion de réunion et un complément d’avis sur des restaurants.

## <a name="how-a-contextual-add-in-displays"></a>Affichage des compléments contextuels

Un complément contextuel activé s’affiche sur une carte, qui est une fenêtre séparée près de l’entité. La carte s’affiche normalement en-dessous de l’entité et le plus centrée possible par rapport à l’entité. S’il n’existe pas suffisamment d’espace en-dessous de l’entité, la carte est placée au-dessus. La capture d’écran suivante illustre l’entité en surbrillance et, dessous, un complément activé (Plans Bing) sur une carte.

**Exemple d’un complément affiché sur une carte**

![Indique une application contextuelle sur une carte](../images/outlook-detected-entity-card.png)

Pour fermer la carte et quitter le complément, il suffit de cliquer n’importe où en dehors de la carte.

## <a name="current-contextual-add-ins"></a>Compléments contextuels actuels

Les compléments contextuels suivants sont installés par défaut pour les utilisateurs qui utilisent des compléments Outlook :

- Plans Bing 
- Réunions suggérées

## <a name="see-also"></a>Voir aussi

- [Complément Outlook : numéro de commande Contoso](https://github.com/OfficeDev/Outlook-Add-In-Contextual-Regex) (exemple de complément contextuel qui est activé en fonction d’une correspondance d’expression régulière)
- [Créer votre premier complément Outlook](../quickstarts/outlook-quickstart.md)
- [Utiliser des règles d’activation d’expression régulière pour afficher un complément Outlook](use-regular-expressions-to-show-an-outlook-add-in.md)
- [Objet Entités](/javascript/api/outlook/office.entities)
