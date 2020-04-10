---
title: Élément PhoneSettings dans le fichier manifeste
description: L’élément PhoneSettings spécifie l’emplacement source et les paramètres de contrôle qui s’appliquent lorsque votre complément de messagerie est utilisé sur un téléphone.
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: cfdf8efc262c1a75142eb06dd71383579598a86f
ms.sourcegitcommit: c6e3bfd3deb77982d0b7082afd6a48678e96e1c3
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/10/2020
ms.locfileid: "43215046"
---
# <a name="phonesettings-element"></a><span data-ttu-id="3f628-103">PhoneSettings, élément</span><span class="sxs-lookup"><span data-stu-id="3f628-103">PhoneSettings element</span></span>

<span data-ttu-id="3f628-104">Spécifie l’emplacement source et les paramètres de contrôle qui s’appliquent lorsque votre complément de messagerie est utilisé sur un téléphone.</span><span class="sxs-lookup"><span data-stu-id="3f628-104">Specifies source location and control settings that apply when your mail add-in is used on a phone.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3f628-105">L' `PhoneSettings` élément est disponible uniquement dans les versions classiques d’Outlook sur le Web (généralement connectées à des versions antérieures du serveur Exchange local) et Outlook 2013 sur Windows.</span><span class="sxs-lookup"><span data-stu-id="3f628-105">The `PhoneSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span> <span data-ttu-id="3f628-106">Pour prendre en charge Outlook sur Android et iOS, reportez-vous à la rubrique [compléments pour Outlook Mobile](../../outlook/outlook-mobile-addins.md).</span><span class="sxs-lookup"><span data-stu-id="3f628-106">To support Outlook on Android and iOS, see [Add-ins for Outlook Mobile](../../outlook/outlook-mobile-addins.md).</span></span>

<span data-ttu-id="3f628-107">**Type de complément :** messagerie</span><span class="sxs-lookup"><span data-stu-id="3f628-107">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="3f628-108">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="3f628-108">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="3f628-109">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="3f628-109">Contained in</span></span>

[<span data-ttu-id="3f628-110">Form</span><span class="sxs-lookup"><span data-stu-id="3f628-110">Form</span></span>](form.md)

