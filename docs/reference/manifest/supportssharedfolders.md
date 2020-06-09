---
title: Élément SupportsSharedFolders dans le fichier manifest
description: L’élément SupportsSharedFolders définit si le complément Outlook est disponible dans les scénarios de délégué.
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 3835f7060cc52a72ff0a5ed4dbdb9f1e09258669
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608711"
---
# <a name="supportssharedfolders-element"></a><span data-ttu-id="8f971-103">Élément SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="8f971-103">SupportsSharedFolders element</span></span>

<span data-ttu-id="8f971-104">Définit si le complément Outlook est disponible dans les scénarios de délégué.</span><span class="sxs-lookup"><span data-stu-id="8f971-104">Defines whether the Outlook add-in is available in delegate scenarios.</span></span> <span data-ttu-id="8f971-105">L’élément **SupportsSharedFolders** est un élément enfant de [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="8f971-105">The **SupportsSharedFolders** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="8f971-106">Ce paramètre est défini sur *false* par défaut.</span><span class="sxs-lookup"><span data-stu-id="8f971-106">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="8f971-107">Seuls Outlook sur le Web et Windows prennent en charge l’élément **SupportsSharedFolders** .</span><span class="sxs-lookup"><span data-stu-id="8f971-107">Only Outlook on the web and Windows support the **SupportsSharedFolders** element.</span></span>
>
> <span data-ttu-id="8f971-108">La prise en charge de cet élément a été introduite dans l’ensemble de conditions requises 1,8.</span><span class="sxs-lookup"><span data-stu-id="8f971-108">Support for this element was introduced in requirement set 1.8.</span></span> <span data-ttu-id="8f971-109">Voir [les clients et les plateformes](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.</span><span class="sxs-lookup"><span data-stu-id="8f971-109">See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

<span data-ttu-id="8f971-110">Voici un exemple de l’élément **SupportsSharedFolders** .</span><span class="sxs-lookup"><span data-stu-id="8f971-110">The following is an example of the **SupportsSharedFolders** element.</span></span>

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
