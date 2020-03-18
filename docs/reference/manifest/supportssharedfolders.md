---
title: Élément SupportsSharedFolders dans le fichier manifest
description: L’élément SupportsSharedFolders définit si le complément Outlook est disponible dans les scénarios de délégué.
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 66a426b0c31bda61feb23cb83d63722898dfb503
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717886"
---
# <a name="supportssharedfolders-element"></a><span data-ttu-id="764b4-103">Élément SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="764b4-103">SupportsSharedFolders element</span></span>

<span data-ttu-id="764b4-104">Définit si le complément Outlook est disponible dans les scénarios de délégué.</span><span class="sxs-lookup"><span data-stu-id="764b4-104">Defines whether the Outlook add-in is available in delegate scenarios.</span></span> <span data-ttu-id="764b4-105">L’élément **SupportsSharedFolders** est un élément enfant de [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="764b4-105">The **SupportsSharedFolders** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="764b4-106">Ce paramètre est défini sur *false* par défaut.</span><span class="sxs-lookup"><span data-stu-id="764b4-106">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="764b4-107">Seuls Outlook sur le Web et Windows prennent en charge l’élément **SupportsSharedFolders** .</span><span class="sxs-lookup"><span data-stu-id="764b4-107">Only Outlook on the web and Windows support the **SupportsSharedFolders** element.</span></span>
>
> <span data-ttu-id="764b4-108">La prise en charge de cet élément a été introduite dans l’ensemble de conditions requises 1,8.</span><span class="sxs-lookup"><span data-stu-id="764b4-108">Support for this element was introduced in requirement set 1.8.</span></span> <span data-ttu-id="764b4-109">Voir [les clients et les plateformes](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.</span><span class="sxs-lookup"><span data-stu-id="764b4-109">See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

<span data-ttu-id="764b4-110">Voici un exemple de l’élément **SupportsSharedFolders** .</span><span class="sxs-lookup"><span data-stu-id="764b4-110">The following is an example of the **SupportsSharedFolders** element.</span></span>

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
