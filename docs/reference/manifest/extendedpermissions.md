---
title: Élément ExtendedPermissions dans le fichier manifeste
description: Définit la collection d’autorisations étendues dont le complément a besoin pour accéder aux API ou fonctionnalités associées.
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: cf59d13d794f8f303da6cc0ca39066584bc3f56c
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611532"
---
# <a name="extendedpermissions-element"></a>Élément ExtendedPermissions

Définit la collection d’autorisations étendues dont le complément a besoin pour accéder aux API ou fonctionnalités associées. L' `ExtendedPermissions` élément est un élément enfant de [VersionOverrides](versionoverrides.md).

> [!IMPORTANT]
> Cet élément est disponible uniquement dans l' [ensemble de conditions requises d’aperçu des compléments Outlook](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) sur Exchange Online. Les compléments qui utilisent cet élément ne peuvent pas être publiés dans AppSource ou déployés via la fonctionnalité de déploiement centralisée.

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Requis  |  Description  |
|:-----|:-----:|:-----|
|  [ExtendedPermission](extendedpermission.md)    |  Non   | Définit une autorisation étendue dont le complément a besoin pour accéder à l’API ou à la fonctionnalité associée. |

## <a name="extendedpermissions-example"></a>`ExtendedPermissions`tels

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
