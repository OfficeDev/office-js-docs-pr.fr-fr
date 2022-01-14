---
title: Élément VersionOverrides 1.0 dans le fichier manifeste d’un module de messagerie
description: Documentation de référence de l’élément VersionOverrides (courrier) Office fichiers manifeste (XML) des modules.
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: b433a52ad922fb3d397993a3861038f2f82ff165
ms.sourcegitcommit: 9b0e70bb296a84adfaea0d6fee54916be9e13031
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/14/2022
ms.locfileid: "62042177"
---
# <a name="versionoverrides-10-element-in-the-manifest-file-for-a-mail-add-in"></a>Élément VersionOverrides 1.0 dans le fichier manifeste d’un module de messagerie

Cet élément contient des informations pour les fonctionnalités qui ne sont pas pris en charge dans le manifeste de base.

> [!NOTE]
> Cet article suppose que vous connaissez la vue d’ensemble de l’élément [VersionOverrides,](versionoverrides.md)qui contient des informations importantes sur les attributs et les variantes de l’élément.

**Type de complément :** messagerie

**Valide uniquement dans ces schémas VersionOverrides**:

- Mail 1.0

Pour plus d’informations, voir [Remplacements de version dans le manifeste.](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)

**Associés à ces ensembles de conditions requises**:

- [Mailbox 1.3](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md)
- Certains éléments enfants peuvent être associés à des ensembles de conditions requises supplémentaires.

## <a name="child-elements"></a>Éléments enfants

Le tableau suivant s’applique uniquement à la version 1.0 des éléments **VersionOverrides** et uniquement aux modules de messagerie.

> [!NOTE]
> Dans iOS, seule `<WebApplicationInfo>` la prise en charge est prise en charge. Tous les autres éléments enfants **de VersionOverrides** sont ignorés.

|  Élément |  Requis  |  Description  |
|:-----|:-----|:-----|
|  [Description](#description)    |  Non   |  Décrit le complément. |
|  [Configuration requise](requirements.md)  |  Non   |  Spécifie les ensembles de conditions requises minimaux qui doivent être pris en charge pour que le marques du parent `VersionOverrides` prenne effet. Cela doit toujours être *plus* restrictif que l’élément `Requirements` dans la partie base du manifeste.|
|  [Hôtes](hosts.md)                |  Oui  |  Spécifie une collection d’applications Office de données. L’élément Hosts enfant remplace l’élément Hosts dans la partie parent du manifeste.  |
|  [Ressources](resources.md)    |  Oui  | Définit une collection de ressources (chaînes, URL et images) qui sont référencées par d’autres éléments de manifeste.|
|  **VersionOverrides**    |  Non  | Définit des commandes de complément sous une version plus récente du schéma. Voir [Mise en œuvre de plusieurs versions](#implementing-multiple-versions) pour plus d’informations. |
|  [WebApplicationInfo](webapplicationinfo.md)    |  Non  | Spécifie des détails sur l’inscription du add-in auprès d’émetteurs de jetons sécurisés, tels que Azure Active Directory V2.0. |

### <a name="description"></a>Description

Décrit le complément. Cela remplace l’élément `Description` dans une partie parent du manifeste. Le texte de la description est contenu dans un élément enfant de l’élément **LongString** contenu dans l’élément [Resources](resources.md). L’attribut de l’élément Description ne peut pas comporter plus de 32 caractères et doit correspondre à la valeur de l’attribut d’un élément enfant de l’élément `resid`  `id` **ShortString** contenu dans l’élément [Resources.](resources.md) 

**Type de add-in :** Volet De tâches, Courrier

**Valide uniquement dans ces schémas VersionOverrides**:

- Volet De tâches 1.0
- Mail 1.0
- Courrier 1.1

Pour plus d’informations, voir [Remplacements de version dans le manifeste.](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)

**Associés à ces ensembles de conditions requises**:

- [AddinCommands 1.1 lorsque](../requirement-sets/add-in-commands-requirement-sets.md) le parent `<VersionOverrides>` est type Taskpane 1.0.
- [Boîte aux lettres 1.3 lorsque](../../reference/objectmodel/requirement-set-1.3/outlook-requirement-set-1.3.md) le parent `<VersionOverrides>` est de type Mail 1.0.
- [Boîte aux lettres 1.5 lorsque](../../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md) le parent `<VersionOverrides>` est de type Mail 1.1.

## <a name="example"></a>Exemple

Voici un exemple simple. Pour obtenir des exemples plus complets, consultez les manifestes des exemples de Office exemples de [code de la version de l’exemple.](https://github.com/OfficeDev/PnP-OfficeAddins)

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
