---
title: Élément Form dans le fichier manifeste
description: ''
ms.date: 01/13/2020
localization_priority: Normal
ms.openlocfilehash: d545d471e007f0077a8310b0b847bbbf99a8f7ac
ms.sourcegitcommit: dc42e0276007f8ab006028b9cd0cc1526c1bd100
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/14/2020
ms.locfileid: "41120648"
---
# <a name="form-element"></a><span data-ttu-id="b9196-102">Form, élément</span><span class="sxs-lookup"><span data-stu-id="b9196-102">Form element</span></span>

<span data-ttu-id="b9196-103">Paramètres UX pour les formulaires que votre complément de messagerie utilisera lors de l’exécution sur un appareil particulier (ordinateur de bureau, tablette ou téléphone).</span><span class="sxs-lookup"><span data-stu-id="b9196-103">UX settings for the forms that your mail add-in will use when running on a particular device (desktop, tablet, or phone).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b9196-104">Les `DesktopSettings`éléments `TabletSettings`, et `PhoneSettings` sont disponibles uniquement dans les versions classiques d’Outlook sur le Web (généralement connectées à des versions plus anciennes de serveur Exchange local) et Outlook 2013 sur Windows.</span><span class="sxs-lookup"><span data-stu-id="b9196-104">The `DesktopSettings`, `TabletSettings`, and `PhoneSettings` elements are available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span>

<span data-ttu-id="b9196-105">**Type de complément :** messagerie</span><span class="sxs-lookup"><span data-stu-id="b9196-105">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="b9196-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="b9196-106">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="b9196-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="b9196-107">Contained in</span></span>

[<span data-ttu-id="b9196-108">FormSettings</span><span class="sxs-lookup"><span data-stu-id="b9196-108">FormSettings</span></span>](formsettings.md)


## <a name="can-contain"></a><span data-ttu-id="b9196-109">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="b9196-109">Can contain</span></span>

|<span data-ttu-id="b9196-110">**Élément**</span><span class="sxs-lookup"><span data-stu-id="b9196-110">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="b9196-111">DesktopSettings</span><span class="sxs-lookup"><span data-stu-id="b9196-111">DesktopSettings</span></span>](desktopsettings.md)|
|[<span data-ttu-id="b9196-112">TabletSettings</span><span class="sxs-lookup"><span data-stu-id="b9196-112">TabletSettings</span></span>](tabletsettings.md)|
|[<span data-ttu-id="b9196-113">PhoneSettings</span><span class="sxs-lookup"><span data-stu-id="b9196-113">PhoneSettings</span></span>](phonesettings.md)|
