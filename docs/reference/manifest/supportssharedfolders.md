---
title: Élément SupportsSharedFolders dans le fichier manifest
description: L’élément SupportsSharedFolders définit si le Outlook est disponible dans les dossiers partagés et les scénarios de boîtes aux lettres partagées.
ms.date: 09/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 8e13393f10b12e0a3c5ca1b004b202eb2970d264
ms.sourcegitcommit: 489befc41e543a4fb3c504fd9b3f61322134c1ef
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/06/2021
ms.locfileid: "60138722"
---
# <a name="supportssharedfolders-element"></a>Élément SupportsSharedFolders

Définit si le Outlook est disponible dans les scénarios de boîte aux lettres partagée (désormais en prévisualisation) et de dossiers partagés (autrement dit, accès délégué). L’élément **SupportsSharedFolders** est un élément enfant de [DesktopFormFactor](desktopformfactor.md). Ce paramètre est défini sur *false* par défaut.

> [!IMPORTANT]
> La prise en charge de cet élément a été introduite dans l’ensemble de conditions requises 1.8. Voir [les clients et les plateformes](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.

**Type de complément :** messagerie

**Valide uniquement dans ces schémas VersionOverrides**:

- Courrier 1.1

Pour plus d’informations, voir [Remplacements de version dans le manifeste.](../../develop/add-in-manifests.md#version-overrides-in-the-manifest)

**Associés à ces ensembles de conditions requises**:

- [Mailbox 1.8](../../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md)

Voici un exemple de **l’élément SupportsSharedFolders.**

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
