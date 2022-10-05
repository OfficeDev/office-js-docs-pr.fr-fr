---
title: Compléments Outlook contextuels
description: Lancer des tâches liées à un message sans laisser le message lui-même pour faciliter et enrichir l'expérience utilisateur.
ms.date: 10/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: 73a13787dac7a6e74db6b919cc01a6dd33d29ab5
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467022"
---
# <a name="contextual-outlook-add-ins"></a>Compléments Outlook contextuels

Contextual add-ins are Outlook add-ins that activate based on text in a message or appointment. By using contextual add-ins, a user can initiate tasks related to a message without leaving the message itself, which results in an easier and richer user experience.

[!include[JSON manifest does not support contextual add-ins](../includes/json-manifest-outlook-contextual-not-supported.md)]

Voici des exemples de compléments contextuels.

- Choix d’une adresse à ouvrir dans un plan du lieu.
- Choix d’une chaîne ouvrant un complément de suggestion de réunion.
- Choisir un numéro de téléphone permet de l’ajouter à vos contacts.


> [!NOTE]
> Les compléments contextuels ne sont pas disponibles actuellement dans Outlook pour Android et iOS. Cette fonctionnalité sera disponible ultérieurement.
>
> La prise en charge de cette fonctionnalité a été introduite dans l’ensemble de conditions requises 1.6. Voir [les clients et les plateformes](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.

## <a name="how-to-make-a-contextual-add-in"></a>Création d’un complément contextuel

Le manifeste d’un complément contextuel doit inclure un élément [ExtensionPoint](/javascript/api/manifest/extensionpoint#detectedentity) avec une attribut `xsi:type` défini sur `DetectedEntity`. Dans l’élément **\<ExtensionPoint\>** , le complément spécifie les entités ou l’expression régulière qui peuvent l’activer. Si une entité est spécifiée, il peut s’agir d’une des propriétés de l’objet [Entités](/javascript/api/outlook/office.entities).

Par conséquent, le manifeste du complément doit contenir un type de règle **ItemHasKnownEntity** ou **Itemhasregularexpressionmatch**. L’exemple suivant montre comment spécifier qu’un complément doit s’activer sur les messages avec une entité détectée qui est un numéro de téléphone.

```XML
<ExtensionPoint xsi:type="DetectedEntity">
  <Label resid="contextLabel" />
  <!--If you opt to include RequestedHeight, it must be between 140px to 450px, inclusive.-->
  <!--<RequestedHeight>360</RequestedHeight>-->
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
> The `EmailAddress` and `Url` entity types do not support highlighting, so they cannot be used to launch a contextual add-in. They can however be combined in a `RuleCollection` rule type as an additional activation criteria.

## <a name="how-to-launch-a-contextual-add-in"></a>Lancement d’un complément contextuel

A user launches a contextual add-in through text, either a known entity or a developer's regular expression. Typically, a user identifies a contextual add-in because the entity is highlighted. The following example shows how highlighting appears in a message. Here the entity (an address) is colored blue and underlined with a dotted blue line. A user launches the contextual add-in by clicking the highlighted entity. 

**Exemple de texte avec l’entité (une adresse) en surbrillance**

![Affiche l’entité mise en surbrillance dans un e-mail.](../images/outlook-detected-entity-highlight.png)
    
Lorsque plusieurs entités ou compléments contextuels sont présents dans un message, l’interaction avec l’utilisateur a lieu comme suit :

- S’il existe plusieurs entités, l’utilisateur doit cliquer sur une autre entité pour lancer le complément de celle-ci.
- Si une entité active plusieurs compléments, chaque complément ouvre un nouvel onglet. L’utilisateur bascule entre les onglets pour passer d’un complément à l’autre. Par exemple, un nom et une adresse peuvent déclencher un complément de téléphone et une carte.
- If a single string contains multiple entities that activate multiple add-ins, the entire string is highlighted, and clicking the string shows all add-ins relevant to the string on separate tabs. For example, a string that describes a proposed meeting at a restaurant might activate the Suggested Meeting add-in and a restaurant rating add-in.

## <a name="how-a-contextual-add-in-displays"></a>Affichage des compléments contextuels

An activated contextual add-in appears in a card, which is a separate window near the entity. The card will normally appear below the entity and centered with respect to the entity as much as possible. If there is not enough room below the entity, the card is placed above it. The following screenshot shows the highlighted entity, and below it, an activated add-in (Bing Maps) in a card.

**Exemple d’un complément affiché sur une carte**

![Montre une application contextuelle dans une carte.](../images/outlook-detected-entity-card.png)

Pour fermer la carte et quitter le complément, il suffit de cliquer n’importe où en dehors de la carte.

## <a name="current-contextual-add-ins"></a>Compléments contextuels actuels

Les compléments contextuels suivants sont installés par défaut pour les utilisateurs disposant de compléments Outlook.

- Plans Bing
- Réunions suggérées

## <a name="see-also"></a>Voir aussi

- [Complément Outlook : numéro de commande Contoso](https://github.com/OfficeDev/Outlook-Add-In-Contextual-Regex) (exemple de complément contextuel qui est activé en fonction d’une correspondance d’expression régulière)
- [Créer votre premier complément Outlook](../quickstarts/outlook-quickstart.md)
- [Utiliser des règles d’activation d’expression régulière pour afficher un complément Outlook](use-regular-expressions-to-show-an-outlook-add-in.md)
- [Objet Entités](/javascript/api/outlook/office.entities)
