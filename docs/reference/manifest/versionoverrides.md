---
title: Élémznr VersionOverrides dans le fichier manifest
description: ''
ms.date: 02/04/2020
localization_priority: Normal
ms.openlocfilehash: 26183caeb4862038d5304607310aa061d37cf3f1
ms.sourcegitcommit: c1dbea577ae6183523fb663d364422d2adbc8bcf
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/05/2020
ms.locfileid: "41773571"
---
# <a name="versionoverrides-element"></a>Élément VersionOverrides

Élément racine qui contient des informations pour les commandes de complément implémentées par le complément. **VersionOverrides** est un élément enfant de l’élément [OfficeApp](./officeapp.md) dans le manifeste. Cet élément est pris en charge dans le schéma de manifeste v1.1 et versions ultérieures, mais est défini dans le schéma VersionOverrides v1.0 ou v1.1.

## <a name="attributes"></a>Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  **xmlns**       |  Oui  |  Espace de noms du schéma VersionOverrides. Les valeurs autorisées varient en fonction de la `<VersionOverrides>` valeur **xsi : type** de cet élément et de la valeur **xsi : type** de `<OfficeApp>` l’élément parent. Voir les [valeurs d’espace de noms](#namespace-values) ci-dessous.|
|  **xsi:type**  |  Oui  | Version du schéma. À ce stade, les seules valeurs valides sont `VersionOverridesV1_0` et `VersionOverridesV1_1`. |

### <a name="namespace-values"></a>Valeurs d’espace de noms

Le code suivant répertorie la valeur requise de la valeur **xmlns** en fonction de la valeur **xsi : type** de `<OfficeApp>` l’élément parent.

- **Taskpaneapp,** prend en charge uniquement la version 1,0 de VersionOverrides **** et le xmlns `http://schemas.microsoft.com/office/taskpaneappversionoverrides`doit être.
- **ContentApp** prend en charge uniquement la version 1,0 de VersionOverrides **** et le xmlns `http://schemas.microsoft.com/office/contentappversionoverrides`doit être.
- **MailApp** prend en charge les versions 1,0 et 1,1 de VersionOverrides, de **** sorte que la valeur de xmlns `<VersionOverrides>` varie en fonction de la valeur **xsi : type** de cet élément :
    - Lorsque **xsi : type** est `VersionOverridesV1_0`, **xmlns** doit être `http://schemas.microsoft.com/office/mailappversionoverrides`.
    - Lorsque **xsi : type** est `VersionOverridesV1_1`, **xmlns** doit être `http://schemas.microsoft.com/office/mailappversionoverrides/1.1`.

> [!NOTE]
> Actuellement, seul Outlook 2016 ou version ultérieure prend en charge le schéma VersionOverrides `VersionOverridesV1_1` v 1.1 et le type.

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Requis  |  Description  |
|:-----|:-----|:-----|
|  **Description**    |  Non   |  Décrit le complément. Cela remplace l’élément `Description` dans une partie parent du manifeste. Le texte de la description est contenu dans un élément enfant de l’élément **LongString** contenu dans l’élément [Resources](./resources.md). L’attribut `resid` de l’élément **Description** est défini sur la valeur de l’attribut `id` de l’élément `String` qui contient le texte.|
|  **Configuration requise**  |  Non   |  Spécifie l’ensemble de conditions requises minimal et la version d’Office.js qui doit être activée par le complément Office. Cela remplace l’élément `Requirements` dans la partie parent du manifeste.|
|  [Hôtes](./hosts.md)                |  Oui  |  Spécifie une collection d’hôtes d’Office. L’élément Hosts enfant remplace l’élément Hosts dans la partie parent du manifeste.  |
|  [Ressources](./resources.md)    |  Oui  | Définit une collection de ressources (chaînes, URL et images) qui sont référencées par d’autres éléments de manifeste.|
|  [EquivalentAddins](./equivalentaddins.md)    |  Non  | Spécifie les compléments natifs (COM/XLL) équivalents au complément Web. Le complément Web n’est pas activé si un complément natif équivalent est installé.|
|  **VersionOverrides**    |  Non  | Définit des commandes de complément sous une version plus récente du schéma. Voir [Mise en œuvre de plusieurs versions](#implementing-multiple-versions) pour plus d’informations. |
|  [WebApplicationInfo](./webapplicationinfo.md)    |  Non  | Fournit des détails sur l’inscription du complément avec des émetteurs de jetons sécurisés, tels qu’Azure Active Directory V 2.0. |

### <a name="versionoverrides-example"></a>Exemple VersionOverrides

Voici un exemple d’un élément typique `<VersionOverrides>` , y compris des éléments enfants qui ne sont pas obligatoires, mais qui sont généralement utilisés.

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

Pour mettre en œuvre à la fois les schémas VersionOverrides v1.0 et v1.1, le manifeste devrait ressembler à l’exemple suivant :

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
