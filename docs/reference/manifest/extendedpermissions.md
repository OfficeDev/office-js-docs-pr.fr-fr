---
title: Élément ExtendedPermissions dans le fichier manifeste
description: Définit la collection d’autorisations étendues dont le complément a besoin pour accéder aux API ou fonctionnalités associées.
ms.date: 10/15/2020
localization_priority: Normal
ms.openlocfilehash: 1e3aa16c160613d34ef2c4f9c25bc2ffe4970816
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/20/2020
ms.locfileid: "48626441"
---
# <a name="extendedpermissions-element"></a>Élément ExtendedPermissions

Définit la collection d’autorisations étendues dont le complément a besoin pour accéder aux API ou fonctionnalités associées. L' `ExtendedPermissions` élément est un élément enfant de [VersionOverrides](versionoverrides.md).

> [!IMPORTANT]
> La prise en charge de cet élément a été introduite dans l’ensemble de conditions requises 1,9. Voir [les clients et les plateformes](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Requis  |  Description  |
|:-----|:-----:|:-----|
|  [ExtendedPermission](extendedpermission.md)    |  Non   | Définit une autorisation étendue dont le complément a besoin pour accéder à l’API ou à la fonctionnalité associée. |

## <a name="extendedpermissions-example"></a>`ExtendedPermissions` tels

Voici un exemple de l' `ExtendedPermissions` élément.

```XML
...
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    ...
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <SupportsSharedFolders>true</SupportsSharedFolders>
          <FunctionFile resid="residDesktopFuncUrl" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <!-- Configure selected extension point. -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed. -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
    ...
    <ExtendedPermissions>
      <ExtendedPermission>AppendOnSend</ExtendedPermission>
    </ExtendedPermissions>
  </VersionOverrides>
</VersionOverrides>
...
```

## <a name="contained-in"></a>Contenu dans

[VersionOverrides](versionoverrides.md)

## <a name="can-contain"></a>Peut contenir

[ExtendedPermission](extendedpermission.md)
