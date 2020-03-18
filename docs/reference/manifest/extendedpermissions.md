---
title: Élément ExtendedPermissions dans le fichier manifeste
description: Définit la collection d’autorisations étendues dont le complément a besoin pour accéder aux API ou fonctionnalités associées.
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: 86d898052af6ba0e6f6bc8b341fff9f0f8408967
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718222"
---
# <a name="extendedpermissions-element"></a><span data-ttu-id="b1781-103">Élément ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="b1781-103">ExtendedPermissions element</span></span>

<span data-ttu-id="b1781-104">Définit la collection d’autorisations étendues dont le complément a besoin pour accéder aux API ou fonctionnalités associées.</span><span class="sxs-lookup"><span data-stu-id="b1781-104">Defines the collection of extended permissions the add-in needs to access associated APIs or features.</span></span> <span data-ttu-id="b1781-105">L' `ExtendedPermissions` élément est un élément enfant de [VersionOverrides](versionoverrides.md).</span><span class="sxs-lookup"><span data-stu-id="b1781-105">The `ExtendedPermissions` element is a child element of [VersionOverrides](versionoverrides.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b1781-106">Cet élément est disponible uniquement dans l' [ensemble de conditions requises d’aperçu des compléments Outlook](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) sur Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="b1781-106">This element is only available in the [Outlook add-ins preview requirement set](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) against Exchange Online.</span></span> <span data-ttu-id="b1781-107">Les compléments qui utilisent cet élément ne peuvent pas être publiés dans AppSource ou déployés via la fonctionnalité de déploiement centralisée.</span><span class="sxs-lookup"><span data-stu-id="b1781-107">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

## <a name="child-elements"></a><span data-ttu-id="b1781-108">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="b1781-108">Child elements</span></span>

|  <span data-ttu-id="b1781-109">Élément</span><span class="sxs-lookup"><span data-stu-id="b1781-109">Element</span></span> |  <span data-ttu-id="b1781-110">Requis</span><span class="sxs-lookup"><span data-stu-id="b1781-110">Required</span></span>  |  <span data-ttu-id="b1781-111">Description</span><span class="sxs-lookup"><span data-stu-id="b1781-111">Description</span></span>  |
|:-----|:-----:|:-----|
|  [<span data-ttu-id="b1781-112">ExtendedPermission</span><span class="sxs-lookup"><span data-stu-id="b1781-112">ExtendedPermission</span></span>](extendedpermission.md)    |  <span data-ttu-id="b1781-113">Non</span><span class="sxs-lookup"><span data-stu-id="b1781-113">No</span></span>   | <span data-ttu-id="b1781-114">Définit une autorisation étendue dont le complément a besoin pour accéder à l’API ou à la fonctionnalité associée.</span><span class="sxs-lookup"><span data-stu-id="b1781-114">Defines an extended permission needed for the add-in to access the associated API or feature.</span></span> |

## <a name="extendedpermissions-example"></a><span data-ttu-id="b1781-115">`ExtendedPermissions`tels</span><span class="sxs-lookup"><span data-stu-id="b1781-115">`ExtendedPermissions` example</span></span>

<span data-ttu-id="b1781-116">Voici un exemple de l' `ExtendedPermissions` élément.</span><span class="sxs-lookup"><span data-stu-id="b1781-116">The following is an example of the `ExtendedPermissions` element.</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="b1781-117">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="b1781-117">Contained in</span></span>

[<span data-ttu-id="b1781-118">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="b1781-118">VersionOverrides</span></span>](versionoverrides.md)

## <a name="can-contain"></a><span data-ttu-id="b1781-119">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="b1781-119">Can contain</span></span>

[<span data-ttu-id="b1781-120">ExtendedPermission</span><span class="sxs-lookup"><span data-stu-id="b1781-120">ExtendedPermission</span></span>](extendedpermission.md)
