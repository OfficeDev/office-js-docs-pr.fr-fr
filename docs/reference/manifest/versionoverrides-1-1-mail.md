---
title: Élément VersionOverrides 1.1 dans le fichier manifeste d’un module de messagerie
description: Documentation de référence de l’élément VersionOverrides 1.1 (messagerie) pour les Office manifeste des modules (XML).
ms.date: 02/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 7e826dad6e4586c83ece8aaa7b083f74b69fade0
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340665"
---
# <a name="versionoverrides-11-element-in-the-manifest-file-for-a-mail-add-in"></a>Élément VersionOverrides 1.1 dans le fichier manifeste d’un module de messagerie

Cet élément contient des informations sur les fonctionnalités qui ne sont pas pris en charge dans le manifeste de base.

> [!NOTE]
> Cet article suppose que vous connaissez la vue d’ensemble de l’élément [VersionOverrides](versionoverrides.md), qui contient des informations importantes sur les attributs et les variantes de l’élément.

**Type de complément :** messagerie

**Valide uniquement dans les schémas VersionOverrides ci-après** :

- Courrier 1.1

Pour plus d’informations, voir [Remplacements de version dans le manifeste](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

**Associés à ces ensembles de conditions requises** :

- [Mailbox 1.5](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)
- Certains éléments enfants peuvent être associés à des ensembles de conditions requises supplémentaires.

## <a name="child-elements"></a>Éléments enfants

Le tableau suivant s’applique uniquement à la version 1.1 des éléments **VersionOverrides** et uniquement aux modules de messagerie.

> [!NOTE]
> Dans iOS, seul **WebApplicationInfo** est pris en charge. Tous les autres éléments enfants **de VersionOverrides** sont ignorés.

|  Élément |  Requis  |  Description  |
|:-----|:-----|:-----|
|  [Description](#description)    |  Non   |  Décrit le complément. |
|  [Configuration requise](requirements.md)  |  Non   |  Spécifie les ensembles de conditions requises minimales qui doivent être pris en charge pour que le markup dans le **parent VersionOverrides** prenne effet. Cela doit toujours être *plus* restrictif que **l’élément Requirements** dans la partie base du manifeste.|
|  [Hôtes](hosts.md)                |  Oui  |  Spécifie une collection d’applications Office de données. L’élément Hosts enfant remplace l’élément Hosts dans la partie parent du manifeste.  |
|  [Ressources](resources.md)    |  Oui  | Définit une collection de ressources (chaînes, URL et images) qui sont référencées par d’autres éléments de manifeste.|
|  [EquivalentAddins](equivalentaddins.md)    |  Non  | Spécifie les compl?ments natifs (COM/XLL) qui sont équivalents au compl?ment web. Le add-in web n’est pas activé si un application native équivalente est installé.|
|  **VersionOverrides**    |  Non  | Actuellement insaisissable dans VersionOverrides 1.1 pour les modules de messagerie. |
|  [WebApplicationInfo](webapplicationinfo.md)    |  Non  | Spécifie des détails sur l’inscription du add-in auprès d’émetteurs de jetons sécurisés, tels que Azure Active Directory V2.0. |
|  [ExtendedPermissions](extendedpermissions.md) |  Non  |  Spécifie une collection d’autorisations étendues. |

### <a name="description"></a>Description

Décrit le complément. Remplace l’élément **Description** dans toute partie parent du manifeste. Le texte de la description est contenu dans un élément enfant de l’élément **LongString** contenu dans l’élément [Resources](resources.md). L’attribut `resid` de l’élément **Description** ne peut pas comporter plus de 32 `id` caractères et doit correspondre à la valeur de l’attribut d’un élément enfant de l’élément **ShortString** contenu dans l’élément [Resources](resources.md) .

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

## <a name="example"></a>Exemple

Voici un exemple simple. Pour obtenir des exemples plus complexes, consultez les manifestes des exemples de Office des [exemples de code de modules.](https://github.com/OfficeDev/PnP-OfficeAddins)

Voici un exemple d’élément **VersionOverrides** classique, y compris certains éléments enfants qui ne sont pas obligatoires mais qui sont généralement utilisés.

```xml
<OfficeApp ... xsi:type="MailApp">
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources>
      <!-- add information on resources -->
    </Resources>
  </VersionOverrides>
...
</OfficeApp>
```

## <a name="implementing-multiple-versions"></a>Mise en œuvre de plusieurs versions

Un manifeste peut implémenter plusieurs versions de l’élément `VersionOverrides` qui prennent en charge différentes versions du schéma VersionOverrides. Cette opération permet éventuellement la prise en charge de nouvelles fonctionnalités dans un schéma plus récent tout en prenant en charge des clients plus anciens qui ne prennent pas en charge les nouvelles fonctionnalités.

Pour mettre en œuvre plusieurs versions, l’élément `VersionOverrides` de la nouvelle version doit être un enfant de l’élément `VersionOverrides` de l’ancienne version. L’élément enfant `VersionOverrides` n’hérite pas des valeurs du parent.

Pour implémenter les schémas VersionOverrides v1.0 et v1.1, le manifeste ressemblerait à l’exemple suivant.

```xml
<OfficeApp ... xsi:type="MailApp">
...
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Description resid="residDescription" />
    <Requirements>
      <!-- add information on requirements -->
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- add information on form factors -->
      </Host>
    </Hosts>
    <Resources>
      <!-- add information on resources -->
    </Resources>

    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
      <Description resid="residDescription" />
      <Requirements>
        <!-- add information on requirements -->
      </Requirements>
      <Hosts>
        <Host xsi:type="MailHost">
          <!-- add information on form factors -->
        </Host>
      </Hosts>
      <Resources>
        <!-- add information on resources -->
      </Resources>
    </VersionOverrides>  
  </VersionOverrides>
...
</OfficeApp>
```
