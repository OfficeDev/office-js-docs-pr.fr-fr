---
title: Élément PhoneSettings dans le fichier manifeste
description: L’élément PhoneSettings spécifie l’emplacement source et les paramètres de contrôle qui s’appliquent lorsque votre complément de messagerie est utilisé sur un téléphone.
ms.date: 01/13/2020
localization_priority: Normal
ms.openlocfilehash: 581a3ae71a58cd05aac52129a6f4395a60c20cef
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720476"
---
# <a name="phonesettings-element"></a><span data-ttu-id="1b061-103">PhoneSettings, élément</span><span class="sxs-lookup"><span data-stu-id="1b061-103">PhoneSettings element</span></span>

<span data-ttu-id="1b061-104">Spécifie l’emplacement source et les paramètres de contrôle qui s’appliquent lorsque votre complément de messagerie est utilisé sur un téléphone.</span><span class="sxs-lookup"><span data-stu-id="1b061-104">Specifies source location and control settings that apply when your mail add-in is used on a phone.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1b061-105">L' `PhoneSettings` élément est disponible uniquement dans les versions classiques d’Outlook sur le Web (généralement connectées à des versions antérieures du serveur Exchange local) et Outlook 2013 sur Windows.</span><span class="sxs-lookup"><span data-stu-id="1b061-105">The `PhoneSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span> <span data-ttu-id="1b061-106">Pour prendre en charge Outlook sur Android et iOS, reportez-vous à la rubrique [compléments pour Outlook Mobile](../../outlook/outlook-mobile-addins.md).</span><span class="sxs-lookup"><span data-stu-id="1b061-106">To support Outlook on Android and iOS, see [Add-ins for Outlook Mobile](../../outlook/outlook-mobile-addins.md).</span></span>

<span data-ttu-id="1b061-107">**Type de complément :** messagerie</span><span class="sxs-lookup"><span data-stu-id="1b061-107">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="1b061-108">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="1b061-108">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="1b061-109">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="1b061-109">Contained in</span></span>

[<span data-ttu-id="1b061-110">Form</span><span class="sxs-lookup"><span data-stu-id="1b061-110">Form</span></span>](form.md)

