---
title: Élément ExtendedPermissions dans le fichier manifeste
description: Définit la collection d’autorisations étendues dont le complément a besoin pour accéder aux API ou fonctionnalités associées.
ms.date: 10/15/2020
localization_priority: Normal
ms.openlocfilehash: 1e3aa16c160613d34ef2c4f9c25bc2ffe4970816
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/20/2020
ms.locfileid: "48626441"
---
# <a name="extendedpermissions-element"></a><span data-ttu-id="66351-103">Élément ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="66351-103">ExtendedPermissions element</span></span>

<span data-ttu-id="66351-104">Définit la collection d’autorisations étendues dont le complément a besoin pour accéder aux API ou fonctionnalités associées.</span><span class="sxs-lookup"><span data-stu-id="66351-104">Defines the collection of extended permissions the add-in needs to access associated APIs or features.</span></span> <span data-ttu-id="66351-105">L' `ExtendedPermissions` élément est un élément enfant de [VersionOverrides](versionoverrides.md).</span><span class="sxs-lookup"><span data-stu-id="66351-105">The `ExtendedPermissions` element is a child element of [VersionOverrides](versionoverrides.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="66351-106">La prise en charge de cet élément a été introduite dans l’ensemble de conditions requises 1,9.</span><span class="sxs-lookup"><span data-stu-id="66351-106">Support for this element was introduced in requirement set 1.9.</span></span> <span data-ttu-id="66351-107">Voir [les clients et les plateformes](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.</span><span class="sxs-lookup"><span data-stu-id="66351-107">See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="child-elements"></a><span data-ttu-id="66351-108">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="66351-108">Child elements</span></span>

|  <span data-ttu-id="66351-109">Élément</span><span class="sxs-lookup"><span data-stu-id="66351-109">Element</span></span> |  <span data-ttu-id="66351-110">Requis</span><span class="sxs-lookup"><span data-stu-id="66351-110">Required</span></span>  |  <span data-ttu-id="66351-111">Description</span><span class="sxs-lookup"><span data-stu-id="66351-111">Description</span></span>  |
|:-----|:-----:|:-----|
|  [<span data-ttu-id="66351-112">ExtendedPermission</span><span class="sxs-lookup"><span data-stu-id="66351-112">ExtendedPermission</span></span>](extendedpermission.md)    |  <span data-ttu-id="66351-113">Non</span><span class="sxs-lookup"><span data-stu-id="66351-113">No</span></span>   | <span data-ttu-id="66351-114">Définit une autorisation étendue dont le complément a besoin pour accéder à l’API ou à la fonctionnalité associée.</span><span class="sxs-lookup"><span data-stu-id="66351-114">Defines an extended permission needed for the add-in to access the associated API or feature.</span></span> |

## <a name="extendedpermissions-example"></a><span data-ttu-id="66351-115">`ExtendedPermissions` tels</span><span class="sxs-lookup"><span data-stu-id="66351-115">`ExtendedPermissions` example</span></span>

<span data-ttu-id="66351-116">Voici un exemple de l' `ExtendedPermissions` élément.</span><span class="sxs-lookup"><span data-stu-id="66351-116">The following is an example of the `ExtendedPermissions` element.</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="66351-117">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="66351-117">Contained in</span></span>

[<span data-ttu-id="66351-118">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="66351-118">VersionOverrides</span></span>](versionoverrides.md)

## <a name="can-contain"></a><span data-ttu-id="66351-119">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="66351-119">Can contain</span></span>

[<span data-ttu-id="66351-120">ExtendedPermission</span><span class="sxs-lookup"><span data-stu-id="66351-120">ExtendedPermission</span></span>](extendedpermission.md)
