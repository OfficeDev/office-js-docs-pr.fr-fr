---
title: Élément de contrôle de type MobileButton dans le fichier manifeste
description: Définit un bouton sur un appareil mobile qui exécute une action ou lance un volet Des tâches.
ms.date: 02/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: d498b728bf7f19cf239ffc6178f19cdf9a62de58
ms.sourcegitcommit: d01aa8101630031515bf27f14361c5a3062c3ec4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/09/2022
ms.locfileid: "62467902"
---
# <a name="control-element-of-type-mobilebutton"></a>Élément Control de type MobileButton

Définit un bouton qui exécute une action ou lance un volet Des tâches et qui apparaît uniquement sur les plateformes mobiles.

> [!NOTE]
> Cet article suppose une connaissance de [l’article](control.md) de référence du contrôle de base qui contient des informations importantes sur les attributs de l’élément.

Un bouton mobile effectue une action unique lorsque l’utilisateur le sélectionne. Il peut exécuter une fonction ou afficher un volet Office. Chaque contrôle de bouton doit avoir une valeur `id` d’attribut unique parmi tous **les éléments Control** du manifeste.

**Type de complément :** messagerie

**Valide uniquement dans les schémas VersionOverrides ci-après** :

- Courrier 1.1

La valeur `MobileButton` de **xsi:type** est définie dans le schéma VersionOverrides 1.1. Pour les éléments [VersionOverrides](versionoverrides.md) la contenant, l’attribut `xsi:type` doit avoir la valeur `VersionOverridesV1_1`.

Pour plus d’informations, voir [Remplacements de version dans le manifeste](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associés à ces ensembles de conditions requises** :

- [Mailbox 1.5](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Requis  |  Description  |
|:-----|:-----|:-----|
|  [Label](#label)     | Oui |  Texte du bouton. |
|  [Icon](icon.md)      | Oui |  Image du bouton.         |
|  [Action](action.md)    | Oui |  Spécifie l’action à effectuer. Il ne peut y avoir **qu’un seul enfant Action** **d’un élément** Control. |

### <a name="label"></a>Étiquette

Spécifie le texte du bouton au moyen de son seul attribut, **resid**, qui ne peut pas être plus de 32 caractères et doit être définie sur la valeur de **l’attribut id** d’un élément **String** dans l’enfant **ShortStrings** de l’élément [Resources](resources.md) .

**Type de complément :** messagerie

**Valide uniquement dans les schémas VersionOverrides ci-après** :

- Courrier 1.1

Pour plus d’informations, voir [Remplacements de version dans le manifeste](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associés à ces ensembles de conditions requises** :

- [Mailbox 1.5](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)

## <a name="examples"></a>Exemples

Dans l’exemple suivant, le bouton exécute une fonction.

```xml
<Control xsi:type="MobileButton" id="Contoso.msgReadFunctionButton">
  <Label resid="funcReadButtonLabel" />
  <Icon>
    <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
    <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
    <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
    <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
    <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
    <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
    <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
    <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
    <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>getSubject</FunctionName>
  </Action>
</Control>
```

Dans l’exemple suivant, le bouton affiche un volet Des tâches.

```xml
<Control xsi:type="MobileButton" id="Contoso.msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Icon>
    <bt:Image resid="blue-icon-16-1" size="25" scale="1" />
    <bt:Image resid="blue-icon-16-2" size="25" scale="2" />
    <bt:Image resid="blue-icon-16-3" size="25" scale="3" />
    <bt:Image resid="blue-icon-32-1" size="32" scale="1" />
    <bt:Image resid="blue-icon-32-2" size="32" scale="2" />
    <bt:Image resid="blue-icon-32-3" size="32" scale="3" />
    <bt:Image resid="blue-icon-80-1" size="48" scale="1" />
    <bt:Image resid="blue-icon-80-2" size="48" scale="2" />
    <bt:Image resid="blue-icon-80-3" size="48" scale="3" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
  </Action>
</Control>
```
