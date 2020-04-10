---
title: Élément Form dans le fichier manifeste
description: Paramètres UX pour les formulaires que votre complément de messagerie utilisera lors de l’exécution sur un appareil particulier (ordinateur de bureau, tablette ou téléphone).
ms.date: 04/09/2020
localization_priority: Normal
ms.openlocfilehash: 3e8d60c13a72a50090075d7cd16a0719498c4982
ms.sourcegitcommit: c6e3bfd3deb77982d0b7082afd6a48678e96e1c3
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/10/2020
ms.locfileid: "43215067"
---
# <a name="form-element"></a><span data-ttu-id="f1f6c-103">Élément Form</span><span class="sxs-lookup"><span data-stu-id="f1f6c-103">Form element</span></span>

<span data-ttu-id="f1f6c-104">Paramètres UX pour les formulaires que votre complément de messagerie utilisera lors de l’exécution sur un appareil particulier (ordinateur de bureau, tablette ou téléphone).</span><span class="sxs-lookup"><span data-stu-id="f1f6c-104">UX settings for the forms that your mail add-in will use when running on a particular device (desktop, tablet, or phone).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="f1f6c-105">Les `DesktopSettings`éléments `TabletSettings`, et `PhoneSettings` sont disponibles uniquement dans les versions classiques d’Outlook sur le Web (généralement connectées à des versions plus anciennes de serveur Exchange local) et Outlook 2013 sur Windows.</span><span class="sxs-lookup"><span data-stu-id="f1f6c-105">The `DesktopSettings`, `TabletSettings`, and `PhoneSettings` elements are available only in classic Outlook on the web (usually connected to older versions of on-premises Exchange server) and Outlook 2013 on Windows.</span></span>

<span data-ttu-id="f1f6c-106">**Type de complément :** messagerie</span><span class="sxs-lookup"><span data-stu-id="f1f6c-106">**Add-in type:** Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="f1f6c-107">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="f1f6c-107">Syntax</span></span>

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

## <a name="contained-in"></a><span data-ttu-id="f1f6c-108">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="f1f6c-108">Contained in</span></span>

[<span data-ttu-id="f1f6c-109">FormSettings</span><span class="sxs-lookup"><span data-stu-id="f1f6c-109">FormSettings</span></span>](formsettings.md)


## <a name="can-contain"></a><span data-ttu-id="f1f6c-110">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="f1f6c-110">Can contain</span></span>

|<span data-ttu-id="f1f6c-111">**Élément**</span><span class="sxs-lookup"><span data-stu-id="f1f6c-111">**Element**</span></span>|
|:-----|
|[<span data-ttu-id="f1f6c-112">DesktopSettings</span><span class="sxs-lookup"><span data-stu-id="f1f6c-112">DesktopSettings</span></span>](desktopsettings.md)|
|[<span data-ttu-id="f1f6c-113">TabletSettings</span><span class="sxs-lookup"><span data-stu-id="f1f6c-113">TabletSettings</span></span>](tabletsettings.md)|
|[<span data-ttu-id="f1f6c-114">PhoneSettings</span><span class="sxs-lookup"><span data-stu-id="f1f6c-114">PhoneSettings</span></span>](phonesettings.md)|
