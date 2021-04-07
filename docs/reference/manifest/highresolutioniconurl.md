---
title: Élément HighResolutionIconUrl dans le fichier manifeste
description: Spécifie l’URL de l’image qui est utilisée pour représenter votre complément Office dans l’UX d’insertion UX et l’Office Store sur les écrans à haute résolution (DPI).
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: 42a7ebf0e02eb365962b574821d5a7004a8b867f
ms.sourcegitcommit: 0bff0411d8cfefd4bb00c189643358e6fb1df95e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/07/2021
ms.locfileid: "51604643"
---
# <a name="highresolutioniconurl-element"></a><span data-ttu-id="61e5c-103">HighResolutionIconUrl, élément</span><span class="sxs-lookup"><span data-stu-id="61e5c-103">HighResolutionIconUrl element</span></span>

<span data-ttu-id="61e5c-104">Spécifie l’URL de l’image qui est utilisée pour représenter votre complément Office dans l’UX d’insertion UX et l’Office Store sur les écrans à haute résolution (DPI).</span><span class="sxs-lookup"><span data-stu-id="61e5c-104">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store on high DPI screens.</span></span>

<span data-ttu-id="61e5c-105">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="61e5c-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="61e5c-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="61e5c-106">Syntax</span></span>

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="61e5c-107">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="61e5c-107">Can contain</span></span>

[<span data-ttu-id="61e5c-108">Override</span><span class="sxs-lookup"><span data-stu-id="61e5c-108">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="61e5c-109">Attributs</span><span class="sxs-lookup"><span data-stu-id="61e5c-109">Attributes</span></span>

|<span data-ttu-id="61e5c-110">Attribut</span><span class="sxs-lookup"><span data-stu-id="61e5c-110">Attribute</span></span>|<span data-ttu-id="61e5c-111">Type</span><span class="sxs-lookup"><span data-stu-id="61e5c-111">Type</span></span>|<span data-ttu-id="61e5c-112">Requis</span><span class="sxs-lookup"><span data-stu-id="61e5c-112">Required</span></span>|<span data-ttu-id="61e5c-113">Description</span><span class="sxs-lookup"><span data-stu-id="61e5c-113">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="61e5c-114">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="61e5c-114">DefaultValue</span></span>|<span data-ttu-id="61e5c-115">chaîne (URL)</span><span class="sxs-lookup"><span data-stu-id="61e5c-115">string (URL)</span></span>|<span data-ttu-id="61e5c-116">obligatoire</span><span class="sxs-lookup"><span data-stu-id="61e5c-116">required</span></span>|<span data-ttu-id="61e5c-117">Spécifie la valeur par défaut de ce paramètre, exprimée pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="61e5c-117">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="61e5c-118">Remarques</span><span class="sxs-lookup"><span data-stu-id="61e5c-118">Remarks</span></span>

<span data-ttu-id="61e5c-119">Pour un module de messagerie, l’icône s’affiche dans l’interface utilisateur gérer les  >  **fichiers des modules.**</span><span class="sxs-lookup"><span data-stu-id="61e5c-119">For a mail add-in, the icon is displayed in the **File** > **Manage add-ins** UI .</span></span> <span data-ttu-id="61e5c-120">Pour un complément de contenu ou de volet Office, l’icône s’affiche dans l’interface utilisateur, sous **Insérer** > **Compléments**.</span><span class="sxs-lookup"><span data-stu-id="61e5c-120">For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span></span>

<span data-ttu-id="61e5c-121">L’image doit être dans un des formats de fichier suivants : GIF, JPG, PNG, EXIF, BMP ou TIFF.</span><span class="sxs-lookup"><span data-stu-id="61e5c-121">The image must be in one of the following file formats: GIF, JPG, PNG, EXIF, BMP, or TIFF.</span></span> <span data-ttu-id="61e5c-122">Pour les applications de contenu et de volet de tâches, la résolution d’image doit être de 64 x 64 pixels.</span><span class="sxs-lookup"><span data-stu-id="61e5c-122">For content and task pane apps, the image resolution must be 64 x 64 pixels.</span></span> <span data-ttu-id="61e5c-123">Pour les applications de messagerie, l’image doit faire 128 x 128 pixels.</span><span class="sxs-lookup"><span data-stu-id="61e5c-123">For mail apps, the image must be 128 x 128 pixels.</span></span> <span data-ttu-id="61e5c-124">Pour plus d’informations, voir la section _Créer une identité visuelle cohérente pour votre application_ dans [Création de listings efficaces dans AppSource et dans Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span><span class="sxs-lookup"><span data-stu-id="61e5c-124">For more information, see the section  _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span></span>
