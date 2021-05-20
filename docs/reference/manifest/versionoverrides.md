---
title: Élémznr VersionOverrides dans le fichier manifest
description: Documentation de référence de l’élément VersionOverrides Office fichiers manifestes add-ins (XML).
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 0a70ded82b4603b1ac70698947a4710a4a44b5b6
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555149"
---
# <a name="versionoverrides-element"></a>Élément VersionOverrides

Élément racine qui contient des informations pour les commandes de complément implémentées par le complément. **VersionOverrides** est un élément enfant de l’élément [OfficeApp](officeapp.md) dans le manifeste. Cet élément est pris en charge dans le schéma de manifeste v1.1 et versions ultérieures, mais est défini dans le schéma VersionOverrides v1.0 ou v1.1.

## <a name="attributes"></a>Attributs

|  Attribut  |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  **xmlns**       |  Oui  |  L’espace nominatif versionOverrides schéma. Les valeurs autorisées varient en fonction de `<VersionOverrides>` la valeur **xsi:type de cet** élément et de la valeur **xsi:type** de l’élément `<OfficeApp>` parent. Voir les [valeurs namespace](#namespace-values) ci-dessous.|
|  **xsi:type**  |  Oui  | Version du schéma. À ce stade, les seules valeurs valides sont `VersionOverridesV1_0` et `VersionOverridesV1_1`. |

### <a name="namespace-values"></a>Valeurs de l’espace nominatif

Ce qui suit énumère la valeur requise de la **valeur xmlns** en fonction de la **valeur xsi:type** de l’élément `<OfficeApp>` parent.

- **TaskPaneApp prend** en charge uniquement la version 1.0 de VersionOverrides, et **les xmlns** doivent être `http://schemas.microsoft.com/office/taskpaneappversionoverrides` .
- **ContentApp** prend en charge uniquement la version 1.0 de VersionOverrides, et **les xmlns** doivent être `http://schemas.microsoft.com/office/contentappversionoverrides` .
- **MailApp** prend en charge les versions 1.0 et 1.1 de VersionOverrides, de sorte que la valeur **des xmlns** varie en fonction de la `<VersionOverrides>` valeur **xsi:type de cet** élément :
    - Lorsque **xsi:type** est `VersionOverridesV1_0` , **xmlns** doit être `http://schemas.microsoft.com/office/mailappversionoverrides` .
    - Lorsque **xsi:type** est `VersionOverridesV1_1` , **xmlns** doit être `http://schemas.microsoft.com/office/mailappversionoverrides/1.1` .

> [!NOTE]
> Actuellement, Outlook 2016 ou plus tard prend en charge le schéma VersionOverrides v1.1 et le `VersionOverridesV1_1` type.

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Obligatoire  |  Description  |
|:-----|:-----|:-----|
|  **Description**    |  Non   |  Décrit le complément. Cela remplace l’élément `Description` dans une partie parent du manifeste. Le texte de la description est contenu dans un élément enfant de l’élément **LongString** contenu dans l’élément [Resources](resources.md). `resid`L’attribut de **l’élément Description** ne peut pas être supérieur à 32 caractères et est défini sur la valeur de `id` `String` l’attribut de l’élément qui contient le texte.|
|  **Configuration requise**  |  Non   |  Spécifie l’ensemble de conditions requises minimal et la version d’Office.js qui doit être activée par le complément Office. Cela remplace l’élément `Requirements` dans la partie parent du manifeste.|
|  [Hôtes](hosts.md)                |  Oui  |  Spécifie une collection d’applications Office’utilisation. L’élément hôtes de l’enfant l’emporte sur l’élément Hôtes dans la partie parente du manifeste.  |
|  [Ressources](resources.md)    |  Oui  | Définit une collection de ressources (chaînes, URL et images) qui sont référencées par d’autres éléments de manifeste.|
|  [EquivalentAddins](equivalentaddins.md)    |  Non  | Spécifie les modules d’ajout natifs (COM/XLL) équivalents à l’add-in web. L’add-in web n’est pas activé si un module d’ajout natif équivalent est installé.|
|  **VersionOverrides**    |  Non  | Définit des commandes de complément sous une version plus récente du schéma. Voir [Mise en œuvre de plusieurs versions](#implementing-multiple-versions) pour plus d’informations. |
|  [WebApplicationInfo](webapplicationinfo.md)    |  Non  | Spécifie les détails de l’enregistrement de l’add-in auprès d’émetteurs de jetons sécurisés, tels que Azure Active Directory V2.0. |
|  [ExtendedPermissions](extendedpermissions.md) |  Non  |  Spécifie une collection d’autorisations étendues. |

### <a name="versionoverrides-example"></a>Exemple VersionOverrides

Ce qui suit est un exemple d’élément `<VersionOverrides>` typique, y compris certains éléments enfant qui ne sont pas nécessaires, mais sont généralement utilisés.

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
