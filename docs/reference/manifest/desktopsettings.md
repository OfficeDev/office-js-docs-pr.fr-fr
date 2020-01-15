---
title: Élément DesktopSettings dans le fichier manifest
description: ''
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 6dfa69d407e267a1cbcfdeaad0bdf9cdf75c1465
ms.sourcegitcommit: dc42e0276007f8ab006028b9cd0cc1526c1bd100
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/14/2020
ms.locfileid: "41120641"
---
# <a name="desktopsettings-element"></a><span data-ttu-id="43c8c-102">Élément DesktopSettings</span><span class="sxs-lookup"><span data-stu-id="43c8c-102">DesktopSettings element</span></span>

<span data-ttu-id="43c8c-103">Spécifie l’emplacement source et les paramètres de contrôle qui s’appliquent lorsque votre complément de messagerie est utilisé sur un ordinateur de bureau.</span><span class="sxs-lookup"><span data-stu-id="43c8c-103">Specifies source location and control settings that apply when your mail add-in is used on a desktop computer.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="43c8c-104">L' `DesktopSettings` élément est disponible uniquement dans les versions classiques d’Outlook sur le Web (généralement connectées à des versions antérieures du serveur Exchange local) et Outlook 2013 sur Windows.</span><span class="sxs-lookup"><span data-stu-id="43c8c-104">The `DesktopSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span>

<span data-ttu-id="43c8c-105">**Type de complément :** messagerie</span><span class="sxs-lookup"><span data-stu-id="43c8c-105">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="43c8c-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="43c8c-106">Syntax</span></span>

```XML
<Form xsi:type="ItemRead">
   <!--website.html is a placeholder for your own add-in website.-->
   <DesktopSettings>
      <SourceLocation DefaultValue="https://website.html" />
      <!--RequestedHeight must be between 240px to 800px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
   </DesktopSettings>
   <TabletSettings>
      <SourceLocation DefaultValue="https://website.html" />
      <!--RequestedHeight must be between 240px to 800px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
   </TabletSettings>
   <PhoneSettings>
      <SourceLocation DefaultValue="https://website.html" />
   </PhoneSettings>
</Form>
```

## <a name="contained-in"></a><span data-ttu-id="43c8c-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="43c8c-107">Contained in</span></span>

[<span data-ttu-id="43c8c-108">Form</span><span class="sxs-lookup"><span data-stu-id="43c8c-108">Form</span></span>](form.md)
