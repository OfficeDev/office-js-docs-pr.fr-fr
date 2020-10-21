---
title: Élément ExtendedPermission dans le fichier manifeste
description: Définit une autorisation étendue dont le complément a besoin pour accéder à l’API ou la fonctionnalité associée.
ms.date: 10/15/2020
localization_priority: Normal
ms.openlocfilehash: 996cac59c44220d05165c7be6ae7c3d79d853271
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/20/2020
ms.locfileid: "48626399"
---
# <a name="extendedpermission-element"></a>`ExtendedPermission` élément

Définit une autorisation étendue dont le complément a besoin pour accéder à l’API ou la fonctionnalité associée. L' `ExtendedPermission` élément est un élément enfant de [ExtendedPermissions](extendedpermissions.md).

> [!IMPORTANT]
> La prise en charge de cet élément a été introduite dans l’ensemble de conditions requises 1,9. Voir [les clients et les plateformes](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.

## <a name="available-extended-permissions"></a>Autorisations étendues disponibles

Les valeurs suivantes sont disponibles.

|Valeur disponible|Description|Hôtes|
|---|---|---|
|`AppendOnSend`|Déclare que le complément utilise l’API [Office. Body. appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendonsendasync-data--options--callback-) .|Outlook|

## <a name="extendedpermission-example"></a>`ExtendedPermission` tels

Voici un exemple de l' `ExtendedPermission` élément.

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

[ExtendedPermissions](extendedpermissions.md)
