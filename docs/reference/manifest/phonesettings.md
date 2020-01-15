---
title: Élément PhoneSettings dans le fichier manifeste
description: ''
ms.date: 01/13/2020
localization_priority: Normal
ms.openlocfilehash: e3ea104af7e634b4e6e6cbeaac395af11ae4e376
ms.sourcegitcommit: dc42e0276007f8ab006028b9cd0cc1526c1bd100
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/14/2020
ms.locfileid: "41120655"
---
# <a name="phonesettings-element"></a><span data-ttu-id="8db2c-102">PhoneSettings, élément</span><span class="sxs-lookup"><span data-stu-id="8db2c-102">PhoneSettings element</span></span>

<span data-ttu-id="8db2c-103">Spécifie l’emplacement source et les paramètres de contrôle qui s’appliquent lorsque votre complément de messagerie est utilisé sur un téléphone.</span><span class="sxs-lookup"><span data-stu-id="8db2c-103">Specifies source location and control settings that apply when your mail add-in is used on a phone.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="8db2c-104">L' `PhoneSettings` élément est disponible uniquement dans les versions classiques d’Outlook sur le Web (généralement connectées à des versions antérieures du serveur Exchange local) et Outlook 2013 sur Windows.</span><span class="sxs-lookup"><span data-stu-id="8db2c-104">The `PhoneSettings` element is available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span> <span data-ttu-id="8db2c-105">Pour prendre en charge Outlook sur Android et iOS, reportez-vous à la rubrique [compléments pour Outlook Mobile](/outlook/add-ins/outlook-mobile-addins).</span><span class="sxs-lookup"><span data-stu-id="8db2c-105">To support Outlook on Android and iOS, see [Add-ins for Outlook Mobile](/outlook/add-ins/outlook-mobile-addins).</span></span>

<span data-ttu-id="8db2c-106">**Type de complément :** messagerie</span><span class="sxs-lookup"><span data-stu-id="8db2c-106">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="8db2c-107">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="8db2c-107">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="8db2c-108">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="8db2c-108">Contained in</span></span>

[<span data-ttu-id="8db2c-109">Form</span><span class="sxs-lookup"><span data-stu-id="8db2c-109">Form</span></span>](form.md)

