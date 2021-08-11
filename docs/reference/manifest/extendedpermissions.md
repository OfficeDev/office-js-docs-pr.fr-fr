---
title: Élément ExtendedPermissions dans le fichier manifeste
description: Définit la collection d’autorisations étendues dont le add-in a besoin pour accéder aux API ou fonctionnalités associées.
ms.date: 10/15/2020
localization_priority: Normal
ms.openlocfilehash: c3f021adfcc2f3a4ba7b7d7aeeb52f3213d92788d401130abbc92618930d09fe
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57097894"
---
# <a name="extendedpermissions-element"></a>Élément ExtendedPermissions

Définit la collection d’autorisations étendues dont le add-in a besoin pour accéder aux API ou fonctionnalités associées. `ExtendedPermissions`L’élément est un élément enfant de [VersionOverrides](versionoverrides.md).

> [!IMPORTANT]
> La prise en charge de cet élément a été introduite dans l’ensemble de conditions requises 1.9. Voir [les clients et les plateformes](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Obligatoire  |  Description  |
|:-----|:-----:|:-----|
|  [ExtendedPermission](extendedpermission.md)    |  Non   | Définit une autorisation étendue nécessaire pour que le add-in accède à l’API ou à la fonctionnalité associée. |

## <a name="extendedpermissions-example"></a>`ExtendedPermissions` exemple

Voici un exemple de `ExtendedPermissions` l’élément.

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
