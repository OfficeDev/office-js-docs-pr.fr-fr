---
title: Élément Form dans le fichier manifeste
description: Paramètres UX pour les formulaires que votre complément de messagerie utilisera lors de l’exécution sur un appareil particulier (ordinateur de bureau, tablette ou téléphone).
ms.date: 01/13/2020
localization_priority: Normal
ms.openlocfilehash: 9b1696b2fecf6b07ee2a3c0a31611d4f2ad1f291
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718208"
---
# <a name="form-element"></a><span data-ttu-id="3a716-103">Élément Form</span><span class="sxs-lookup"><span data-stu-id="3a716-103">Form element</span></span>

<span data-ttu-id="3a716-104">Paramètres UX pour les formulaires que votre complément de messagerie utilisera lors de l’exécution sur un appareil particulier (ordinateur de bureau, tablette ou téléphone).</span><span class="sxs-lookup"><span data-stu-id="3a716-104">UX settings for the forms that your mail add-in will use when running on a particular device (desktop, tablet, or phone).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3a716-105">Les `DesktopSettings`éléments `TabletSettings`, et `PhoneSettings` sont disponibles uniquement dans les versions classiques d’Outlook sur le Web (généralement connectées à des versions plus anciennes de serveur Exchange local) et Outlook 2013 sur Windows.</span><span class="sxs-lookup"><span data-stu-id="3a716-105">The `DesktopSettings`, `TabletSettings`, and `PhoneSettings` elements are available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span>

<span data-ttu-id="3a716-106">**Type de complément :** messagerie</span><span class="sxs-lookup"><span data-stu-id="3a716-106">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="3a716-107">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="3a716-107">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="3a716-108">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="3a716-108">Contained in</span></span>

[<span data-ttu-id="3a716-109">FormSettings</span><span class="sxs-lookup"><span data-stu-id="3a716-109">FormSettings</span></span>](formsettings.md)


## <a name="can-contain"></a><span data-ttu-id="3a716-110">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="3a716-110">Can contain</span></span>

|<span data-ttu-id="3a716-111">**Élément**</span><span class="sxs-lookup"><span data-stu-id="3a716-111">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="3a716-112">DesktopSettings</span><span class="sxs-lookup"><span data-stu-id="3a716-112">DesktopSettings</span></span>](desktopsettings.md)|
|[<span data-ttu-id="3a716-113">TabletSettings</span><span class="sxs-lookup"><span data-stu-id="3a716-113">TabletSettings</span></span>](tabletsettings.md)|
|[<span data-ttu-id="3a716-114">PhoneSettings</span><span class="sxs-lookup"><span data-stu-id="3a716-114">PhoneSettings</span></span>](phonesettings.md)|
