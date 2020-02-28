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
# <a name="supportssharedfolders-element"></a><span data-ttu-id="64ea0-102">Élément SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="64ea0-102">SupportsSharedFolders element</span></span>

<span data-ttu-id="64ea0-103">Définit si le complément Outlook est disponible dans les scénarios de délégué.</span><span class="sxs-lookup"><span data-stu-id="64ea0-103">Defines whether the Outlook add-in is available in delegate scenarios.</span></span> <span data-ttu-id="64ea0-104">L’élément **SupportsSharedFolders** est un élément enfant de [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="64ea0-104">The **SupportsSharedFolders** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="64ea0-105">Ce paramètre est défini sur *false* par défaut.</span><span class="sxs-lookup"><span data-stu-id="64ea0-105">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="64ea0-106">Seuls Outlook sur le Web et Windows prennent en charge l’élément **SupportsSharedFolders** .</span><span class="sxs-lookup"><span data-stu-id="64ea0-106">Only Outlook on the web and Windows support the **SupportsSharedFolders** element.</span></span>
>
> <span data-ttu-id="64ea0-107">La prise en charge de cet élément a été introduite dans l’ensemble de conditions requises 1,8.</span><span class="sxs-lookup"><span data-stu-id="64ea0-107">Support for this element was introduced in requirement set 1.8.</span></span> <span data-ttu-id="64ea0-108">Voir [les clients et les plateformes](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.</span><span class="sxs-lookup"><span data-stu-id="64ea0-108">See [clients and platforms](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

<span data-ttu-id="64ea0-109">Voici un exemple de l’élément **SupportsSharedFolders** .</span><span class="sxs-lookup"><span data-stu-id="64ea0-109">The following is an example of the **SupportsSharedFolders** element.</span></span>

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
