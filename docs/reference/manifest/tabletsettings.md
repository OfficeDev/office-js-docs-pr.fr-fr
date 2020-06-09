---
title: Élément TabletSettings dans le fichier manifeste
description: L’élément TabletSettings spécifie les paramètres de contrôle qui s’appliquent lorsque votre complément de messagerie est utilisé sur une tablette.
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: b5a74db4f9fb43df10a08ab43b59507f6e0d7952
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608697"
---
# <a name="tabletsettings-element"></a><span data-ttu-id="ea0a5-103">TabletSettings, élément</span><span class="sxs-lookup"><span data-stu-id="ea0a5-103">TabletSettings element</span></span>

<span data-ttu-id="ea0a5-104">Spécifie les paramètres de contrôle qui s’appliquent lorsque votre complément de messagerie est utilisé sur une tablette.</span><span class="sxs-lookup"><span data-stu-id="ea0a5-104">Specifies control settings that apply when your mail add-in is used on a tablet.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="ea0a5-105">L' `TabletSettings` élément est disponible uniquement dans les versions classiques d’Outlook sur le Web (généralement connectées à des versions antérieures du serveur Exchange local) et outlook 2013 sur Windows.</span><span class="sxs-lookup"><span data-stu-id="ea0a5-105">The `TabletSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span> <span data-ttu-id="ea0a5-106">Pour prendre en charge Outlook sur Android et iOS, reportez-vous à la rubrique [compléments pour Outlook Mobile](../../outlook/outlook-mobile-addins.md).</span><span class="sxs-lookup"><span data-stu-id="ea0a5-106">To support Outlook on Android and iOS, see [Add-ins for Outlook Mobile](../../outlook/outlook-mobile-addins.md).</span></span>

<span data-ttu-id="ea0a5-107">**Type de complément :** messagerie</span><span class="sxs-lookup"><span data-stu-id="ea0a5-107">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="ea0a5-108">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="ea0a5-108">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="ea0a5-109">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="ea0a5-109">Contained in</span></span>

[<span data-ttu-id="ea0a5-110">Form</span><span class="sxs-lookup"><span data-stu-id="ea0a5-110">Form</span></span>](form.md)
