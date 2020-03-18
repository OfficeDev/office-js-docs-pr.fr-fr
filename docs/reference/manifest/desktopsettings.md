---
title: Élément DesktopSettings dans le fichier manifest
description: Spécifie l’emplacement source et les paramètres de contrôle qui s’appliquent lorsque votre complément de messagerie est utilisé sur un ordinateur de bureau.
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: d48532482fc71fec2a96133ee8e813cae798613f
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718355"
---
# <a name="desktopsettings-element"></a><span data-ttu-id="48157-103">DesktopSettings, élément</span><span class="sxs-lookup"><span data-stu-id="48157-103">DesktopSettings element</span></span>

<span data-ttu-id="48157-104">Spécifie l’emplacement source et les paramètres de contrôle qui s’appliquent lorsque votre complément de messagerie est utilisé sur un ordinateur de bureau.</span><span class="sxs-lookup"><span data-stu-id="48157-104">Specifies source location and control settings that apply when your mail add-in is used on a desktop computer.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="48157-105">L' `DesktopSettings` élément est disponible uniquement dans les versions classiques d’Outlook sur le Web (généralement connectées à des versions antérieures du serveur Exchange local) et Outlook 2013 sur Windows.</span><span class="sxs-lookup"><span data-stu-id="48157-105">The `DesktopSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span>

<span data-ttu-id="48157-106">**Type de complément :** messagerie</span><span class="sxs-lookup"><span data-stu-id="48157-106">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="48157-107">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="48157-107">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="48157-108">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="48157-108">Contained in</span></span>

[<span data-ttu-id="48157-109">Form</span><span class="sxs-lookup"><span data-stu-id="48157-109">Form</span></span>](form.md)
