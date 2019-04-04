---
title: Élément SupportsSharedFolders dans le fichier manifest
description: ''
ms.date: 04/02/2019
localization_priority: Normal
ms.openlocfilehash: 976f8ba00f6ac9ac32def56933af1077527b7e9c
ms.sourcegitcommit: cb763661c927a1c7ec03feeda92a343537ad7fba
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/03/2019
ms.locfileid: "31396904"
---
# <a name="supportssharedfolders-element"></a>Élément SupportsSharedFolders

Définit si le complément Outlook est disponible dans les scénarios de délégué. L’élément **SupportsSharedFolders** est un élément enfant de [DesktopFormFactor](desktopformfactor.md). Ce paramètre est défini sur *false* par défaut.

> [!IMPORTANT]
> L'accès délégué pour les compléments Outlook est actuellement [en](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview) préversion et uniquement pris en charge dans les clients qui s'exécutent sur Exchange Online. Les compléments qui utilisent cet élément ne peuvent pas être publiés dans AppSource ou déployés via la fonctionnalité déploiement centralisée.

L’exemple suivant présente l’élément**SupportsSharedFolders**.

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
            <!-- configure selected extension point -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
    ...
  </VersionOverrides>
</VersionOverrides>
...
```
