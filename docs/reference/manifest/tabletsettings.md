---
title: Élément TabletSettings dans le fichier manifeste
description: L’élément TabletSettings spécifie les paramètres de contrôle qui s’appliquent lorsque votre complément de messagerie est utilisé sur une tablette.
ms.date: 01/13/2020
localization_priority: Normal
ms.openlocfilehash: 2b8b372d27274d89d3aed4b5bacb9faa4893fda5
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717858"
---
# <a name="tabletsettings-element"></a><span data-ttu-id="6baf5-103">TabletSettings, élément</span><span class="sxs-lookup"><span data-stu-id="6baf5-103">TabletSettings element</span></span>

<span data-ttu-id="6baf5-104">Spécifie les paramètres de contrôle qui s’appliquent lorsque votre complément de messagerie est utilisé sur une tablette.</span><span class="sxs-lookup"><span data-stu-id="6baf5-104">Specifies control settings that apply when your mail add-in is used on a tablet.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="6baf5-105">L' `TabletSettings` élément est disponible uniquement dans les versions classiques d’Outlook sur le Web (généralement connectées à des versions antérieures du serveur Exchange local) et Outlook 2013 sur Windows.</span><span class="sxs-lookup"><span data-stu-id="6baf5-105">The `TabletSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span> <span data-ttu-id="6baf5-106">Pour prendre en charge Outlook sur Android et iOS, reportez-vous à la rubrique [compléments pour Outlook Mobile](../../outlook/outlook-mobile-addins.md).</span><span class="sxs-lookup"><span data-stu-id="6baf5-106">To support Outlook on Android and iOS, see [Add-ins for Outlook Mobile](../../outlook/outlook-mobile-addins.md).</span></span>

<span data-ttu-id="6baf5-107">**Type de complément :** messagerie</span><span class="sxs-lookup"><span data-stu-id="6baf5-107">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="6baf5-108">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="6baf5-108">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="6baf5-109">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="6baf5-109">Contained in</span></span>

[<span data-ttu-id="6baf5-110">Form</span><span class="sxs-lookup"><span data-stu-id="6baf5-110">Form</span></span>](form.md)

