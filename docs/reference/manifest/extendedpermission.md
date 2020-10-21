---
title: Élément ExtendedPermission dans le fichier manifeste
description: Définit une autorisation étendue dont le complément a besoin pour accéder à l’API ou la fonctionnalité associée.
ms.date: 10/15/2020
localization_priority: Normal
ms.openlocfilehash: 996cac59c44220d05165c7be6ae7c3d79d853271
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/20/2020
ms.locfileid: "48626399"
---
# <a name="extendedpermission-element"></a><span data-ttu-id="886ac-103">`ExtendedPermission` élément</span><span class="sxs-lookup"><span data-stu-id="886ac-103">`ExtendedPermission` element</span></span>

<span data-ttu-id="886ac-104">Définit une autorisation étendue dont le complément a besoin pour accéder à l’API ou la fonctionnalité associée.</span><span class="sxs-lookup"><span data-stu-id="886ac-104">Defines an extended permission the add-in needs to access the associated API or feature.</span></span> <span data-ttu-id="886ac-105">L' `ExtendedPermission` élément est un élément enfant de [ExtendedPermissions](extendedpermissions.md).</span><span class="sxs-lookup"><span data-stu-id="886ac-105">The `ExtendedPermission` element is a child element of [ExtendedPermissions](extendedpermissions.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="886ac-106">La prise en charge de cet élément a été introduite dans l’ensemble de conditions requises 1,9.</span><span class="sxs-lookup"><span data-stu-id="886ac-106">Support for this element was introduced in requirement set 1.9.</span></span> <span data-ttu-id="886ac-107">Voir [les clients et les plateformes](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) qui prennent en charge cet ensemble de conditions requises.</span><span class="sxs-lookup"><span data-stu-id="886ac-107">See [clients and platforms](../../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.</span></span>

## <a name="available-extended-permissions"></a><span data-ttu-id="886ac-108">Autorisations étendues disponibles</span><span class="sxs-lookup"><span data-stu-id="886ac-108">Available extended permissions</span></span>

<span data-ttu-id="886ac-109">Les valeurs suivantes sont disponibles.</span><span class="sxs-lookup"><span data-stu-id="886ac-109">The following are the available values.</span></span>

|<span data-ttu-id="886ac-110">Valeur disponible</span><span class="sxs-lookup"><span data-stu-id="886ac-110">Available value</span></span>|<span data-ttu-id="886ac-111">Description</span><span class="sxs-lookup"><span data-stu-id="886ac-111">Description</span></span>|<span data-ttu-id="886ac-112">Hôtes</span><span class="sxs-lookup"><span data-stu-id="886ac-112">Hosts</span></span>|
|---|---|---|
|`AppendOnSend`|<span data-ttu-id="886ac-113">Déclare que le complément utilise l’API [Office. Body. appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendonsendasync-data--options--callback-) .</span><span class="sxs-lookup"><span data-stu-id="886ac-113">Declares that the add-in is using the [Office.Body.appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendonsendasync-data--options--callback-) API.</span></span>|<span data-ttu-id="886ac-114">Outlook</span><span class="sxs-lookup"><span data-stu-id="886ac-114">Outlook</span></span>|

## <a name="extendedpermission-example"></a><span data-ttu-id="886ac-115">`ExtendedPermission` tels</span><span class="sxs-lookup"><span data-stu-id="886ac-115">`ExtendedPermission` example</span></span>

<span data-ttu-id="886ac-116">Voici un exemple de l' `ExtendedPermission` élément.</span><span class="sxs-lookup"><span data-stu-id="886ac-116">The following is an example of the `ExtendedPermission` element.</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="886ac-117">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="886ac-117">Contained in</span></span>

[<span data-ttu-id="886ac-118">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="886ac-118">ExtendedPermissions</span></span>](extendedpermissions.md)
