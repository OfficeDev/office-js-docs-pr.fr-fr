---
title: Élément ExtendedPermission dans le fichier manifeste
description: Définit une autorisation étendue dont le complément a besoin pour accéder à l’API ou la fonctionnalité associée.
ms.date: 03/05/2020
localization_priority: Normal
ms.openlocfilehash: 7ff17312ae487d20f4d7af0ed4405cedd8820253
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720602"
---
# <a name="extendedpermission-element"></a><span data-ttu-id="8c755-103">`ExtendedPermission`élément</span><span class="sxs-lookup"><span data-stu-id="8c755-103">`ExtendedPermission` element</span></span>

<span data-ttu-id="8c755-104">Définit une autorisation étendue dont le complément a besoin pour accéder à l’API ou la fonctionnalité associée.</span><span class="sxs-lookup"><span data-stu-id="8c755-104">Defines an extended permission the add-in needs to access the associated API or feature.</span></span> <span data-ttu-id="8c755-105">L' `ExtendedPermission` élément est un élément enfant de [ExtendedPermissions](extendedpermissions.md).</span><span class="sxs-lookup"><span data-stu-id="8c755-105">The `ExtendedPermission` element is a child element of [ExtendedPermissions](extendedpermissions.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="8c755-106">Cet élément est disponible uniquement dans l' [ensemble de conditions requises d’aperçu des compléments Outlook](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) sur Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="8c755-106">This element is only available in the [Outlook add-ins preview requirement set](../objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) against Exchange Online.</span></span> <span data-ttu-id="8c755-107">Les compléments qui utilisent cet élément ne peuvent pas être publiés dans AppSource ou déployés via la fonctionnalité de déploiement centralisée.</span><span class="sxs-lookup"><span data-stu-id="8c755-107">Add-ins that use this element cannot be published to AppSource or deployed via centralized deployment.</span></span>

## <a name="available-extended-permissions"></a><span data-ttu-id="8c755-108">Autorisations étendues disponibles</span><span class="sxs-lookup"><span data-stu-id="8c755-108">Available extended permissions</span></span>

<span data-ttu-id="8c755-109">Les valeurs suivantes sont disponibles.</span><span class="sxs-lookup"><span data-stu-id="8c755-109">The following are the available values.</span></span>

|<span data-ttu-id="8c755-110">Valeur disponible</span><span class="sxs-lookup"><span data-stu-id="8c755-110">Available value</span></span>|<span data-ttu-id="8c755-111">Description</span><span class="sxs-lookup"><span data-stu-id="8c755-111">Description</span></span>|<span data-ttu-id="8c755-112">Hôtes</span><span class="sxs-lookup"><span data-stu-id="8c755-112">Hosts</span></span>|
|---|---|---|
|`AppendOnSend`|<span data-ttu-id="8c755-113">Déclare que le complément utilise l’API [Office. Body. appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-) .</span><span class="sxs-lookup"><span data-stu-id="8c755-113">Declares that the add-in is using the [Office.Body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-) API.</span></span>|<span data-ttu-id="8c755-114">Outlook</span><span class="sxs-lookup"><span data-stu-id="8c755-114">Outlook</span></span>|

## <a name="extendedpermission-example"></a><span data-ttu-id="8c755-115">`ExtendedPermission`tels</span><span class="sxs-lookup"><span data-stu-id="8c755-115">`ExtendedPermission` example</span></span>

<span data-ttu-id="8c755-116">Voici un exemple de l' `ExtendedPermission` élément.</span><span class="sxs-lookup"><span data-stu-id="8c755-116">The following is an example of the `ExtendedPermission` element.</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="8c755-117">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="8c755-117">Contained in</span></span>

[<span data-ttu-id="8c755-118">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="8c755-118">ExtendedPermissions</span></span>](extendedpermissions.md)
