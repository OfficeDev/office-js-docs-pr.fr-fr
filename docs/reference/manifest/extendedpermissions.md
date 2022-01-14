---
title: Élément ExtendedPermissions dans le fichier manifeste
description: Définit la collection d’autorisations étendues dont le add-in a besoin pour accéder aux API ou fonctionnalités associées.
ms.date: 01/04/2022
ms.localizationpriority: medium
ms.openlocfilehash: 46ca6e3e2fb992755d9067b4251200073f07ade1
ms.sourcegitcommit: 9b0e70bb296a84adfaea0d6fee54916be9e13031
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/14/2022
ms.locfileid: "62042125"
---
# <a name="extendedpermissions-element"></a>Élément ExtendedPermissions

Définit la collection d’autorisations étendues dont le add-in a besoin pour accéder aux API ou fonctionnalités associées. `ExtendedPermissions`L’élément est un élément enfant de [VersionOverrides](versionoverrides.md).

> [!IMPORTANT]
> La prise en charge de cet élément a été introduite dans l’ensemble de conditions requises 1.9. Voir [les clients et les plateformes](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.

**Type de complément :** messagerie

**Valide uniquement dans ces schémas VersionOverrides**:

- Courrier 1.1

Pour plus d’informations, voir [Remplacements de version dans le manifeste.](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)

**Associés à ces ensembles de conditions requises**:

- [Mailbox 1.9](../../reference/objectmodel/requirement-set-1.9/outlook-requirement-set-1.9.md)

## <a name="child-elements"></a>Éléments enfants

|  Élément |  Requis  |  Description  |
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
