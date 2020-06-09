---
title: Élément ExtendedPermissions dans le fichier manifeste
description: Définit la collection d’autorisations étendues dont le complément a besoin pour accéder aux API ou fonctionnalités associées.
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: cf59d13d794f8f303da6cc0ca39066584bc3f56c
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611532"
---
# <a name="extendedpermissions-element"></a><span data-ttu-id="29103-103">Élément ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="29103-103">ExtendedPermissions element</span></span>

<span data-ttu-id="29103-104">Définit la collection d’autorisations étendues dont le complément a besoin pour accéder aux API ou fonctionnalités associées.</span><span class="sxs-lookup"><span data-stu-id="29103-104">Defines the collection of extended permissions the add-in needs to access associated APIs or features.</span></span> <span data-ttu-id="29103-105">L' `ExtendedPermissions` élément est un élément enfant de [VersionOverrides](versionoverrides.md).</span><span class="sxs-lookup"><span data-stu-id="29103-105">The `ExtendedPermissions` element is a child element of [VersionOverrides](versionoverrides.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="29103-106">Cet élément est disponible uniquement dans l' [ensemble de conditions requises d’aperçu des compléments Outlook](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) sur Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="29103-106">This element is only available in the [Outlook add-ins preview requirement set](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) against Exchange Online.</span></span> <span data-ttu-id="29103-107">Les compléments qui utilisent cet élément ne peuvent pas être publiés dans AppSource ou déployés via la fonctionnalité de déploiement centralisée.</span><span class="sxs-lookup"><span data-stu-id="29103-107">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

## <a name="child-elements"></a><span data-ttu-id="29103-108">Éléments enfants</span><span class="sxs-lookup"><span data-stu-id="29103-108">Child elements</span></span>

|  <span data-ttu-id="29103-109">Élément</span><span class="sxs-lookup"><span data-stu-id="29103-109">Element</span></span> |  <span data-ttu-id="29103-110">Requis</span><span class="sxs-lookup"><span data-stu-id="29103-110">Required</span></span>  |  <span data-ttu-id="29103-111">Description</span><span class="sxs-lookup"><span data-stu-id="29103-111">Description</span></span>  |
|:-----|:-----:|:-----|
|  [<span data-ttu-id="29103-112">ExtendedPermission</span><span class="sxs-lookup"><span data-stu-id="29103-112">ExtendedPermission</span></span>](extendedpermission.md)    |  <span data-ttu-id="29103-113">Non</span><span class="sxs-lookup"><span data-stu-id="29103-113">No</span></span>   | <span data-ttu-id="29103-114">Définit une autorisation étendue dont le complément a besoin pour accéder à l’API ou à la fonctionnalité associée.</span><span class="sxs-lookup"><span data-stu-id="29103-114">Defines an extended permission needed for the add-in to access the associated API or feature.</span></span> |

## <a name="extendedpermissions-example"></a><span data-ttu-id="29103-115">`ExtendedPermissions`tels</span><span class="sxs-lookup"><span data-stu-id="29103-115">`ExtendedPermissions` example</span></span>

<span data-ttu-id="29103-116">Voici un exemple de l' `ExtendedPermissions` élément.</span><span class="sxs-lookup"><span data-stu-id="29103-116">The following is an example of the `ExtendedPermissions` element.</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="29103-117">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="29103-117">Contained in</span></span>

[<span data-ttu-id="29103-118">VersionOverrides</span><span class="sxs-lookup"><span data-stu-id="29103-118">VersionOverrides</span></span>](versionoverrides.md)

## <a name="can-contain"></a><span data-ttu-id="29103-119">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="29103-119">Can contain</span></span>

[<span data-ttu-id="29103-120">ExtendedPermission</span><span class="sxs-lookup"><span data-stu-id="29103-120">ExtendedPermission</span></span>](extendedpermission.md)
