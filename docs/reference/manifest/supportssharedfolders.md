---
title: Élément SupportsSharedFolders dans le fichier manifest
description: ''
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 81401b79f4c443305e376df7a66a07d916393d17
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596752"
---
# <a name="supportssharedfolders-element"></a><span data-ttu-id="bc419-102">Élément SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="bc419-102">SupportsSharedFolders element</span></span>

<span data-ttu-id="bc419-103">Définit si le complément Outlook est disponible dans les scénarios de délégué.</span><span class="sxs-lookup"><span data-stu-id="bc419-103">Defines whether the Outlook add-in is available in delegate scenarios.</span></span> <span data-ttu-id="bc419-104">L’élément **SupportsSharedFolders** est un élément enfant de [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="bc419-104">The **SupportsSharedFolders** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="bc419-105">Ce paramètre est défini sur *false* par défaut.</span><span class="sxs-lookup"><span data-stu-id="bc419-105">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="bc419-106">Seuls Outlook sur le Web et Windows prennent en charge l’élément **SupportsSharedFolders** .</span><span class="sxs-lookup"><span data-stu-id="bc419-106">Only Outlook on the web and Windows support the **SupportsSharedFolders** element.</span></span>
>
> <span data-ttu-id="bc419-107">La prise en charge de cet élément a été introduite dans l’ensemble de conditions requises 1,8.</span><span class="sxs-lookup"><span data-stu-id="bc419-107">Support for this element was introduced in requirement set 1.8.</span></span> <span data-ttu-id="bc419-108">Voir [les clients et les plateformes](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.</span><span class="sxs-lookup"><span data-stu-id="bc419-108">See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

<span data-ttu-id="bc419-109">Voici un exemple de l’élément **SupportsSharedFolders** .</span><span class="sxs-lookup"><span data-stu-id="bc419-109">The following is an example of the **SupportsSharedFolders** element.</span></span>

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
