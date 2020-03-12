---
title: Élément ExtendedPermissions dans le fichier manifeste
description: ''
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: 966378b8bbed66960d7a99c4a82df75ace1c9161
ms.sourcegitcommit: a0262ea40cd23f221e69bcb0223110f011265d13
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/12/2020
ms.locfileid: "42605805"
---
# <a name="extendedpermissions-element"></a><span data-ttu-id="664ac-102">Élément ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="664ac-102">ExtendedPermissions element</span></span>

<span data-ttu-id="664ac-103">Définit la collection d’autorisations étendues dont le complément a besoin pour accéder aux API ou fonctionnalités associées.</span><span class="sxs-lookup"><span data-stu-id="664ac-103">Defines the collection of extended permissions the add-in needs to access associated APIs or features.</span></span> <span data-ttu-id="664ac-104">L' `ExtendedPermissions` élément est un élément enfant de [VersionOverrides](versionoverrides.md).</span><span class="sxs-lookup"><span data-stu-id="664ac-104">The `ExtendedPermissions` element is a child element of [VersionOverrides](versionoverrides.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="664ac-105">Cet élément est disponible uniquement dans l' [ensemble de conditions requises d’aperçu des compléments Outlook](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) sur Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="664ac-105">This element is only available in the [Outlook add-ins preview requirement set](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) against Exchange Online.</span></span> <span data-ttu-id="664ac-106">Les compléments qui utilisent cet élément ne peuvent pas être publiés dans AppSource ou déployés via la fonctionnalité de déploiement centralisée.</span><span class="sxs-lookup"><span data-stu-id="664ac-106">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

## <a name="child-elements"></a><span data-ttu-id="664ac-107">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="664ac-107">Child elements</span></span>

|  <span data-ttu-id="664ac-108">Élément</span><span class="sxs-lookup"><span data-stu-id="664ac-108">Element</span></span> |  <span data-ttu-id="664ac-109">Requis</span><span class="sxs-lookup"><span data-stu-id="664ac-109">Required</span></span>  |  <span data-ttu-id="664ac-110">Description</span><span class="sxs-lookup"><span data-stu-id="664ac-110">Description</span></span>  |
|:-----|:-----:|:-----|
|  [<span data-ttu-id="664ac-111">ExtendedPermission</span><span class="sxs-lookup"><span data-stu-id="664ac-111">ExtendedPermission</span></span>](extendedpermission.md)    |  <span data-ttu-id="664ac-112">Non</span><span class="sxs-lookup"><span data-stu-id="664ac-112">No</span></span>   | <span data-ttu-id="664ac-113">Définit une autorisation étendue dont le complément a besoin pour accéder à l’API ou à la fonctionnalité associée.</span><span class="sxs-lookup"><span data-stu-id="664ac-113">Defines an extended permission needed for the add-in to access the associated API or feature.</span></span> |

## <a name="extendedpermissions-example"></a><span data-ttu-id="664ac-114">`ExtendedPermissions`tels</span><span class="sxs-lookup"><span data-stu-id="664ac-114">`ExtendedPermissions` example</span></span>

<span data-ttu-id="664ac-115">Voici un exemple de l' `ExtendedPermissions` élément.</span><span class="sxs-lookup"><span data-stu-id="664ac-115">The following is an example of the `ExtendedPermissions` element.</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="664ac-116">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="664ac-116">Contained in</span></span>

[<span data-ttu-id="664ac-117">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="664ac-117">VersionOverrides</span></span>](versionoverrides.md)

## <a name="can-contain"></a><span data-ttu-id="664ac-118">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="664ac-118">Can contain</span></span>

[<span data-ttu-id="664ac-119">ExtendedPermission</span><span class="sxs-lookup"><span data-stu-id="664ac-119">ExtendedPermission</span></span>](extendedpermission.md)
