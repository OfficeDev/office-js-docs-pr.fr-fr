---
title: Élément ExtendedPermission dans le fichier manifeste
description: Définit une autorisation étendue dont le add-in a besoin pour accéder à l’API ou à la fonctionnalité associée.
ms.date: 10/15/2020
localization_priority: Normal
ms.openlocfilehash: 37859350cfaffdc14ab91d5026d67aa0a736ac56
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671757"
---
# <a name="extendedpermission-element"></a>`ExtendedPermission` élément

Définit une autorisation étendue dont le add-in a besoin pour accéder à l’API ou à la fonctionnalité associée. `ExtendedPermission`L’élément est un élément enfant de [ExtendedPermissions](extendedpermissions.md).

> [!IMPORTANT]
> La prise en charge de cet élément a été introduite dans l’ensemble de conditions requises 1.9. Voir [les clients et les plateformes](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.

## <a name="available-extended-permissions"></a>Autorisations étendues disponibles

Voici les valeurs disponibles.

|Valeur disponible|Description|Hôtes|
|---|---|---|
|`AppendOnSend`|Déclare que le add-in utilise le [Office. API Body.appendOnSendAsync.](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendOnSendAsync_data__options__callback_)|Outlook|

## <a name="extendedpermission-example"></a>`ExtendedPermission` exemple

Voici un exemple de `ExtendedPermission` l’élément.

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
