---
title: Élémznr VersionOverrides dans le fichier manifest
description: Documentation de référence de l’élément VersionOverrides pour les fichiers manifeste des compléments Office (XML).
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: 979a75c3ea8b4d600a2c43fc4edfcb0d4e96930e
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431541"
---
# <a name="versionoverrides-element"></a>Élément VersionOverrides

Élément racine qui contient des informations pour les commandes de complément implémentées par le complément. **VersionOverrides** est un élément enfant de l’élément [OfficeApp](./officeapp.md) dans le manifeste. Cet élément est pris en charge dans le schéma de manifeste v1.1 et versions ultérieures, mais est défini dans le schéma VersionOverrides v1.0 ou v1.1.

## <a name="attributes"></a>Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  **xmlns**       |  Oui  |  Espace de noms du schéma VersionOverrides. Les valeurs autorisées varient en fonction de la `<VersionOverrides>` valeur **xsi : type** de cet élément et de la valeur **xsi : type** de l' `<OfficeApp>` élément parent. Voir les [valeurs d’espace de noms](#namespace-values) ci-dessous.|
|  **xsi:type**  |  Oui  | Version du schéma. À ce stade, les seules valeurs valides sont `VersionOverridesV1_0` et `VersionOverridesV1_1`. |

### <a name="namespace-values"></a>Valeurs d’espace de noms

Le code suivant répertorie la valeur requise de la valeur **xmlns** en fonction de la valeur **xsi : type** de l' `<OfficeApp>` élément parent.

- **Taskpaneapp,** prend en charge uniquement la version 1,0 de VersionOverrides et le **xmlns** doit être `http://schemas.microsoft.com/office/taskpaneappversionoverrides` .
- **ContentApp** prend en charge uniquement la version 1,0 de VersionOverrides et le **xmlns** doit être `http://schemas.microsoft.com/office/contentappversionoverrides` .
- **MailApp** prend en charge les versions 1,0 et 1,1 de VersionOverrides, de sorte que la valeur de **xmlns** varie en fonction de la `<VersionOverrides>` valeur **xsi : type** de cet élément :
    - Lorsque **xsi : type** est `VersionOverridesV1_0` , **xmlns** doit être `http://schemas.microsoft.com/office/mailappversionoverrides` .
    - Lorsque **xsi : type** est `VersionOverridesV1_1` , **xmlns** doit être `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` .

> [!NOTE]
> Actuellement, seul Outlook 2016 ou version ultérieure prend en charge le schéma VersionOverrides v 1.1 et le `VersionOverridesV1_1` type.

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  **Description**    |  Non   |  Décrit le complément. Cela remplace l’élément `Description` dans une partie parent du manifeste. Le texte de la description est contenu dans un élément enfant de l’élément **LongString** contenu dans l’élément [Resources](resources.md). L’attribut `resid` de l’élément **Description** est défini sur la valeur de l’attribut `id` de l’élément `String` qui contient le texte.|
|  **Configuration requise**  |  Non   |  Spécifie l’ensemble de conditions requises minimal et la version d’Office.js qui doit être activée par le complément Office. Cela remplace l’élément `Requirements` dans la partie parent du manifeste.|
|  [Hôtes](hosts.md)                |  Oui  |  Spécifie une collection d’applications Office. L’élément hosts enfant remplace l’élément hosts dans la partie parent du manifeste.  |
|  [Ressources](resources.md)    |  Oui  | Définit une collection de ressources (chaînes, URL et images) qui sont référencées par d’autres éléments de manifeste.|
|  [EquivalentAddins](equivalentaddins.md)    |  Non  | Spécifie les compléments natifs (COM/XLL) équivalents au complément Web. Le complément Web n’est pas activé si un complément natif équivalent est installé.|
|  **VersionOverrides**    |  Non  | Définit des commandes de complément sous une version plus récente du schéma. Voir [Mise en œuvre de plusieurs versions](#implementing-multiple-versions) pour plus d’informations. |
|  [WebApplicationInfo](webapplicationinfo.md)    |  Non  | Fournit des détails sur l’inscription du complément avec des émetteurs de jetons sécurisés, tels qu’Azure Active Directory V 2.0. |
|  [ExtendedPermissions](extendedpermissions.md) |  Non  |  Spécifie une collection d’autorisations étendues.<br><br>**Important**: étant donné que l’API [Office. Body. appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendonsendasync-data--options--callback-) est actuellement en préversion, les compléments qui utilisent l' `ExtendedPermissions` élément ne peuvent pas être publiés sur AppSource ou déployés via un déploiement centralisé. |

### <a name="versionoverrides-example"></a>Exemple VersionOverrides

Voici un exemple d’un `<VersionOverrides>` élément typique, y compris des éléments enfants qui ne sont pas obligatoires, mais qui sont généralement utilisés.

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
