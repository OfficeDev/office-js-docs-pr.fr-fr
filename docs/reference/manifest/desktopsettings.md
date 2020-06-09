---
title: Élément DesktopSettings dans le fichier manifest
description: Spécifie l’emplacement source et les paramètres de contrôle qui s’appliquent lorsque votre complément de messagerie est utilisé sur un ordinateur de bureau.
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: 50201080d8be3c8943d16730c34a4bac236d7b90
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612275"
---
# <a name="desktopsettings-element"></a><span data-ttu-id="025d1-103">DesktopSettings, élément</span><span class="sxs-lookup"><span data-stu-id="025d1-103">DesktopSettings element</span></span>

<span data-ttu-id="025d1-104">Spécifie l’emplacement source et les paramètres de contrôle qui s’appliquent lorsque votre complément de messagerie est utilisé sur un ordinateur de bureau.</span><span class="sxs-lookup"><span data-stu-id="025d1-104">Specifies source location and control settings that apply when your mail add-in is used on a desktop computer.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="025d1-105">L' `DesktopSettings` élément est disponible uniquement dans les versions classiques d’Outlook sur le Web (généralement connectées à des versions antérieures du serveur Exchange local) et outlook 2013 sur Windows.</span><span class="sxs-lookup"><span data-stu-id="025d1-105">The `DesktopSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span>

<span data-ttu-id="025d1-106">**Type de complément :** messagerie</span><span class="sxs-lookup"><span data-stu-id="025d1-106">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="025d1-107">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="025d1-107">Syntax</span></span>

```XML
<Form xsi:type="ItemRead">
   <!--https://MyDomain.com/website.html is a placeholder for your own add-in website.-->
   <DesktopSettings>
      <!--If you opt to include RequestedHeight, it must be between 32px to 450px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
      <SourceLocation DefaultValue="https://MyDomain.com/website.html" />
   </DesktopSettings>
   <TabletSettings>
      <!--If you opt to include RequestedHeight, it must be between 32px to 450px, inclusive.-->
      <RequestedHeight>360</RequestedHeight>
      <SourceLocation DefaultValue="https://MyDomain.com/website.html" />
   </TabletSettings>
   <PhoneSettings>
      <SourceLocation DefaultValue="https://MyDomain.com/website.html" />
   </PhoneSettings>
</Form>
```

## <a name="contained-in"></a><span data-ttu-id="025d1-108">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="025d1-108">Contained in</span></span>

[<span data-ttu-id="025d1-109">Form</span><span class="sxs-lookup"><span data-stu-id="025d1-109">Form</span></span>](form.md)
