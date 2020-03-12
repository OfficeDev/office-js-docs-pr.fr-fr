---
title: Élément ExtendedPermission dans le fichier manifeste
description: ''
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: 6c41684fc922f5845559250311edd8182788cfc5
ms.sourcegitcommit: a0262ea40cd23f221e69bcb0223110f011265d13
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/12/2020
ms.locfileid: "42605802"
---
# <a name="extendedpermission-element"></a><span data-ttu-id="29107-102">`ExtendedPermission`élément</span><span class="sxs-lookup"><span data-stu-id="29107-102">`ExtendedPermission` element</span></span>

<span data-ttu-id="29107-103">Définit une autorisation étendue dont le complément a besoin pour accéder à l’API ou la fonctionnalité associée.</span><span class="sxs-lookup"><span data-stu-id="29107-103">Defines an extended permission the add-in needs to access the associated API or feature.</span></span> <span data-ttu-id="29107-104">L' `ExtendedPermission` élément est un élément enfant de [ExtendedPermissions](extendedpermissions.md).</span><span class="sxs-lookup"><span data-stu-id="29107-104">The `ExtendedPermission` element is a child element of [ExtendedPermissions](extendedpermissions.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="29107-105">Cet élément est disponible uniquement dans l' [ensemble de conditions requises d’aperçu des compléments Outlook](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) sur Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="29107-105">This element is only available in the [Outlook add-ins preview requirement set](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) against Exchange Online.</span></span> <span data-ttu-id="29107-106">Les compléments qui utilisent cet élément ne peuvent pas être publiés dans AppSource ou déployés via la fonctionnalité de déploiement centralisée.</span><span class="sxs-lookup"><span data-stu-id="29107-106">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

## <a name="available-extended-permissions"></a><span data-ttu-id="29107-107">Autorisations étendues disponibles</span><span class="sxs-lookup"><span data-stu-id="29107-107">Available extended permissions</span></span>

<span data-ttu-id="29107-108">Les valeurs suivantes sont disponibles.</span><span class="sxs-lookup"><span data-stu-id="29107-108">The following are the available values.</span></span>

|<span data-ttu-id="29107-109">Valeur disponible</span><span class="sxs-lookup"><span data-stu-id="29107-109">Available value</span></span>|<span data-ttu-id="29107-110">Description</span><span class="sxs-lookup"><span data-stu-id="29107-110">Description</span></span>|<span data-ttu-id="29107-111">Hôtes</span><span class="sxs-lookup"><span data-stu-id="29107-111">Hosts</span></span>|
|---|---|---|
|`AppendOnSend`|<span data-ttu-id="29107-112">Déclare que le complément utilise l’API [Office. Body. appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-) .</span><span class="sxs-lookup"><span data-stu-id="29107-112">Declares that the add-in is using the [Office.Body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-) API.</span></span>|<span data-ttu-id="29107-113">Outlook</span><span class="sxs-lookup"><span data-stu-id="29107-113">Outlook</span></span>|

## <a name="extendedpermission-example"></a><span data-ttu-id="29107-114">`ExtendedPermission`tels</span><span class="sxs-lookup"><span data-stu-id="29107-114">`ExtendedPermission` example</span></span>

<span data-ttu-id="29107-115">Voici un exemple de l' `ExtendedPermission` élément.</span><span class="sxs-lookup"><span data-stu-id="29107-115">The following is an example of the `ExtendedPermission` element.</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="29107-116">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="29107-116">Contained in</span></span>

[<span data-ttu-id="29107-117">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="29107-117">ExtendedPermissions</span></span>](extendedpermissions.md)
