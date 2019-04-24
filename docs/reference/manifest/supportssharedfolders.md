---
title: Élément SupportsSharedFolders dans le fichier manifest
description: ''
ms.date: 04/02/2019
localization_priority: Normal
ms.openlocfilehash: 976f8ba00f6ac9ac32def56933af1077527b7e9c
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452038"
---
# <a name="supportssharedfolders-element"></a><span data-ttu-id="c05ec-102">Élément SupportsSharedFolders</span><span class="sxs-lookup"><span data-stu-id="c05ec-102">SupportsSharedFolders element</span></span>

<span data-ttu-id="c05ec-103">Définit si le complément Outlook est disponible dans les scénarios de délégué.</span><span class="sxs-lookup"><span data-stu-id="c05ec-103">Defines whether the Outlook add-in is available in delegate scenarios.</span></span> <span data-ttu-id="c05ec-104">L’élément **SupportsSharedFolders** est un élément enfant de [DesktopFormFactor](desktopformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="c05ec-104">The **SupportsSharedFolders** element is a child element of [DesktopFormFactor](desktopformfactor.md).</span></span> <span data-ttu-id="c05ec-105">Ce paramètre est défini sur *false* par défaut.</span><span class="sxs-lookup"><span data-stu-id="c05ec-105">It is set to *false* by default.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c05ec-106">L'accès délégué pour les compléments Outlook est actuellement [en](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview) préversion et uniquement pris en charge dans les clients qui s'exécutent sur Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="c05ec-106">Delegate access for Outlook add-ins is currently [in preview](/office/dev/add-ins/reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview) and only supported in clients that run against Exchange Online.</span></span> <span data-ttu-id="c05ec-107">Les compléments qui utilisent cet élément ne peuvent pas être publiés dans AppSource ou déployés via la fonctionnalité déploiement centralisée.</span><span class="sxs-lookup"><span data-stu-id="c05ec-107">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

<span data-ttu-id="c05ec-108">L’exemple suivant présente l’élément**SupportsSharedFolders**.</span><span class="sxs-lookup"><span data-stu-id="c05ec-108">The following is an example of the  **SupportsSharedFolders** element.</span></span>

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
