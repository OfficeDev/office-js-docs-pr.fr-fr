---
title: Élément IconUrl dans le fichier manifeste
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: f7eda7ec9e4c5da8ad0b19e5e10649696d4e85c1
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452108"
---
# <a name="iconurl-element"></a><span data-ttu-id="cde1d-102">IconUrl, élément</span><span class="sxs-lookup"><span data-stu-id="cde1d-102">IconUrl element</span></span>

<span data-ttu-id="cde1d-103">Spécifie l’URL de l’image utilisée pour représenter votre complément Office dans l’UX d’insertion UX et l’Office Store.</span><span class="sxs-lookup"><span data-stu-id="cde1d-103">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store.</span></span>

<span data-ttu-id="cde1d-104">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="cde1d-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="cde1d-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="cde1d-105">Syntax</span></span>

```XML
<IconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="cde1d-106">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="cde1d-106">Can contain</span></span>

[<span data-ttu-id="cde1d-107">Override</span><span class="sxs-lookup"><span data-stu-id="cde1d-107">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="cde1d-108">Attributs</span><span class="sxs-lookup"><span data-stu-id="cde1d-108">Attributes</span></span>

|<span data-ttu-id="cde1d-109">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="cde1d-109">**Attribute**</span></span>|<span data-ttu-id="cde1d-110">**Type**</span><span class="sxs-lookup"><span data-stu-id="cde1d-110">**Type**</span></span>|<span data-ttu-id="cde1d-111">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="cde1d-111">**Required**</span></span>|<span data-ttu-id="cde1d-112">**Description**</span><span class="sxs-lookup"><span data-stu-id="cde1d-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="cde1d-113">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="cde1d-113">DefaultValue</span></span>|<span data-ttu-id="cde1d-114">chaîne</span><span class="sxs-lookup"><span data-stu-id="cde1d-114">string</span></span>|<span data-ttu-id="cde1d-115">obligatoire</span><span class="sxs-lookup"><span data-stu-id="cde1d-115">required</span></span>|<span data-ttu-id="cde1d-116">Spécifie la valeur par défaut de ce paramètre, exprimée pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="cde1d-116">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="cde1d-117">Remarques</span><span class="sxs-lookup"><span data-stu-id="cde1d-117">Remarks</span></span>

<span data-ttu-id="cde1d-p101">Pour un complément de messagerie, l’icône s’affiche dans l’interface utilisateur, sous **Fichier**  >  **Gérer les compléments** (Outlook) ou sous **Paramètres**  >  **Gérer les compléments** (Outlook Web App). Pour un complément de contenu ou de volet Office, l’icône s’affiche dans l’interface utilisateur, sous **Insérer**  >  **Compléments**. Pour tous les types de compléments, l’icône est également utilisée sur le site de l’Office Store si vous publiez votre complément dans l’Office Store.</span><span class="sxs-lookup"><span data-stu-id="cde1d-p101">For a mail add-in, the icon is displayed in the  **File** > **Manage add-ins** UI (Outlook) or **Settings** > **Manage add-ins** UI (Outlook Web App). For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI. For all add-in types, the icon is also used on the Office Store site, if you publish your add-in to the Office Store.</span></span>

<span data-ttu-id="cde1d-121">L’image doit être dans un des formats de fichier suivants : GIF, JPG, PNG, EXIF, BMP ou TIFF.</span><span class="sxs-lookup"><span data-stu-id="cde1d-121">The image must be in one of the following file formats: GIF, JPG, PNG, EXIF, BMP, or TIFF.</span></span> <span data-ttu-id="cde1d-122">Pour les applications de volet de tâches et de contenu, l’image spécifiée doit contenir 32 x 32 pixels.</span><span class="sxs-lookup"><span data-stu-id="cde1d-122">For content and task pane apps, the image specified must be 32 x 32 pixels.</span></span> <span data-ttu-id="cde1d-123">Pour les applications de messagerie, la résolution d’image recommandée est de 64 x 64 pixels.</span><span class="sxs-lookup"><span data-stu-id="cde1d-123">For mail apps, the recommended image resolution is 64 x 64 pixels.</span></span> <span data-ttu-id="cde1d-124">Vous devez également spécifier une icône pour une utilisation avec les applications hôte Office en cours d’exécution sur des écrans haute résolution (DPI) à l’aide de l’élément [HighResolutionIconUrl](highresolutioniconurl.md).</span><span class="sxs-lookup"><span data-stu-id="cde1d-124">You should also specify an icon for use with Office host applications running on high DPI screens using the [HighResolutionIconUrl](highresolutioniconurl.md) element.</span></span> <span data-ttu-id="cde1d-125">Pour plus d’informations, reportez-vous à la section _Créer une identité visuelle cohérente pour votre application_ dans [Création de listings efficaces dans AppSource et dans Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span><span class="sxs-lookup"><span data-stu-id="cde1d-125">For more information, see the section _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span></span>
