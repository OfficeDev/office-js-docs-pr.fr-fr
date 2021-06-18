---
title: Élément SupportsSharedFolders dans le fichier manifest
description: L’élément SupportsSharedFolders définit si le Outlook est disponible dans les dossiers partagés et les scénarios de boîtes aux lettres partagées.
ms.date: 06/15/2021
localization_priority: Normal
ms.openlocfilehash: 43f2c60664a6822b714023246cfa044e179e9a55
ms.sourcegitcommit: 0bf0e076f705af29193abe3dba98cbfcce17b24f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/18/2021
ms.locfileid: "53007782"
---
# <a name="supportssharedfolders-element"></a><span data-ttu-id="e92a9-103">Élément SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="e92a9-103">SupportsSharedFolders element</span></span>

<span data-ttu-id="e92a9-104">Définit si le Outlook est disponible dans les scénarios de boîte aux lettres partagée (désormais en prévisualisation) et de dossiers partagés (autrement dit, accès délégué).</span><span class="sxs-lookup"><span data-stu-id="e92a9-104">Defines whether the Outlook add-in is available in shared mailbox (now in preview) and shared folders (that is, delegate access) scenarios.</span></span> <span data-ttu-id="e92a9-105">L’élément **SupportsSharedFolders** est un élément enfant de [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="e92a9-105">The **SupportsSharedFolders** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="e92a9-106">Ce paramètre est défini sur *false* par défaut.</span><span class="sxs-lookup"><span data-stu-id="e92a9-106">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="e92a9-107">La prise en charge de cet élément a été introduite dans l’ensemble de conditions requises 1.8.</span><span class="sxs-lookup"><span data-stu-id="e92a9-107">Support for this element was introduced in requirement set 1.8.</span></span> <span data-ttu-id="e92a9-108">Voir [les clients et les plateformes](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.</span><span class="sxs-lookup"><span data-stu-id="e92a9-108">See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

<span data-ttu-id="e92a9-109">Voici un exemple de **l’élément SupportsSharedFolders.**</span><span class="sxs-lookup"><span data-stu-id="e92a9-109">The following is an example of the **SupportsSharedFolders** element.</span></span>

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
