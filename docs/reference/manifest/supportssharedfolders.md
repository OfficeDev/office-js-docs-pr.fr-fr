---
title: Élément SupportsSharedFolders dans le fichier manifest
description: ''
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: e76d17b618e2aaf15724f15ee6695a932172bba3
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325226"
---
# <a name="supportssharedfolders-element"></a>Élément SupportsSharedFolders

Définit si le complément Outlook est disponible dans les scénarios de délégué. L’élément **SupportsSharedFolders** est un élément enfant de [DesktopFormFactor](desktopformfactor.md). Ce paramètre est défini sur *false* par défaut.

> [!IMPORTANT]
> Seuls Outlook sur le Web et Windows prennent en charge l’élément **SupportsSharedFolders** .
>
> La prise en charge de cet élément a été introduite dans l’ensemble de conditions requises 1,8. Voir [les clients et les plateformes](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.

Voici un exemple de l’élément **SupportsSharedFolders** .

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
  </VersionOverrides>
</VersionOverrides>
...
```
