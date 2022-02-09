---
title: Élément Control de type Button dans le fichier manifeste
description: Définit un bouton qui exécute une action ou lance un volet Des tâches.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: adc58424fe9898bffcbd9e16bed8f3b13b9df4a2
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467891"
---
# <a name="control-element-of-type-button"></a>Élément Control de type Button

Définit un bouton qui exécute une action ou lance un volet Des tâches.

> [!NOTE]
> Cet article suppose une connaissance de [l’article](control.md) de référence du contrôle de base qui contient des informations importantes sur les attributs de l’élément.

Un bouton effectue une action unique quand il est sélectionné. Il peut exécuter une fonction ou afficher un volet Office. Chaque contrôle de bouton doit avoir une valeur `id` d’attribut unique parmi tous **les éléments Control** du manifeste.

> [!IMPORTANT]
> Les contrôles de type « Bouton » sont ignorés sur les plateformes mobiles. Pour prendre en charge les plateformes mobiles, vous devez également avoir un contrôle de type « MobileButton » pour chaque contrôle de type « Button ».

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Requis  |  Description  |
|:-----|:-----|:-----|
|  [Label](#label)     | Oui |  Texte du bouton. |
|  **ToolTip**    |Non|Info-bulle pour le bouton. **L’attribut resid** ne peut pas être plus de 32 caractères et doit être définie sur la valeur de **l’attribut id** d’un **élément String**. **String** est un enfant de l’élément **LongStrings**, qui est lui-même un enfant de l’élément [Resources](resources.md).|
|  [Supertip](supertip.md)  | Oui |  Info-bulle pour le bouton.    |
|  [Icon](icon.md)      | Oui |  Image du bouton.         |
|  [Action](action.md)    | Oui |  Spécifie l’action à effectuer. Il ne peut y avoir **qu’un seul enfant Action** **d’un élément** Control. |
|  [Enabled](enabled.md)    | Non |  Spécifie si le contrôle est activé au lancement du module.  |
|  [OverriddenByRibbonApi](overriddenbyribbonapi.md)      | Non |  Spécifie si le bouton doit apparaître sur les combinaisons d’applications et de plateformes qui prendre en charge les onglets contextuels personnalisés. S’il est utilisé, il doit s’agit du *premier* élément enfant. |

### <a name="label"></a>Étiquette

Spécifie le texte du bouton au moyen de son seul attribut, **resid**, qui ne peut pas être plus de 32 caractères et doit être définie sur la valeur de **l’attribut id** d’un élément **String** dans l’enfant **ShortStrings** de l’élément [Resources](resources.md) .

**Type de add-in :** Volet De tâches, Courrier

**Valide uniquement dans les schémas VersionOverrides ci-après** :

- Volet De tâches 1.0
- Mail 1.0
- Courrier 1.1

Pour plus d’informations, voir [Remplacements de version dans le manifeste](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associés à ces ensembles de conditions requises** :

- [AddinCommands 1.1](../requirement-sets/add-in-commands-requirement-sets.md) lorsque le parent **VersionOverrides** est de type Taskpane 1.0.
- [Boîte aux lettres 1.3 lorsque](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) le parent **VersionOverrides** est de type Mail 1.0.
- [Boîte aux lettres 1.5 lorsque](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) le parent **VersionOverrides** est de type Mail 1.1.

## <a name="examples"></a>Exemples

Dans l’exemple suivant, le bouton exécute une fonction. Il est également configuré pour être désactivé au lancement du module. Il peut être activé par programme. Pour plus d’informations, reportez-vous aux [Commandes Activé et Désactivé pour les compléments](../../design/disable-add-in-commands.md).

```xml
<Control xsi:type="Button" id="Contoso.msgReadFunctionButton">
  <OverriddenByRibbonApi>true</OverriddenByRibbonApi>
  <Label resid="funcReadButtonLabel" />
  <Supertip>
    <Title resid="funcReadSuperTipTitle" />
    <Description resid="funcReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="blue-icon-16" />
    <bt:Image size="32" resid="blue-icon-32" />
    <bt:Image size="80" resid="blue-icon-80" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
  </Action>
  <Enabled>false</Enabled>
</Control>
```

Dans l’exemple suivant, le bouton affiche un volet Des tâches.

```xml
<Control xsi:type="Button" id="Contoso.msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Supertip>
    <Title resid="paneReadSuperTipTitle" />
    <Description resid="paneReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="green-icon-16" />
    <bt:Image size="32" resid="green-icon-32" />
    <bt:Image size="80" resid="green-icon-80" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
</Control>
```
