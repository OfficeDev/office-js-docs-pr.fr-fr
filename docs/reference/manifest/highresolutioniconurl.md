---
title: Élément HighResolutionIconUrl dans le fichier manifeste
description: Spécifie l’URL de l’image qui est utilisée pour représenter votre complément Office dans l’UX d’insertion UX et l’Office Store sur les écrans à haute résolution (DPI).
ms.date: 12/04/2018
localization_priority: Normal
ms.openlocfilehash: 78a9296f38a688073e516fb78a77bb4cdac822c4
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718138"
---
# <a name="highresolutioniconurl-element"></a><span data-ttu-id="44493-103">HighResolutionIconUrl, élément</span><span class="sxs-lookup"><span data-stu-id="44493-103">HighResolutionIconUrl element</span></span>

<span data-ttu-id="44493-104">Spécifie l’URL de l’image qui est utilisée pour représenter votre complément Office dans l’UX d’insertion UX et l’Office Store sur les écrans à haute résolution (DPI).</span><span class="sxs-lookup"><span data-stu-id="44493-104">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store on high DPI screens.</span></span>

<span data-ttu-id="44493-105">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="44493-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="44493-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="44493-106">Syntax</span></span>

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="44493-107">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="44493-107">Can contain</span></span>

[<span data-ttu-id="44493-108">Override</span><span class="sxs-lookup"><span data-stu-id="44493-108">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="44493-109">Attributs</span><span class="sxs-lookup"><span data-stu-id="44493-109">Attributes</span></span>

|<span data-ttu-id="44493-110">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="44493-110">**Attribute**</span></span>|<span data-ttu-id="44493-111">**Type**</span><span class="sxs-lookup"><span data-stu-id="44493-111">**Type**</span></span>|<span data-ttu-id="44493-112">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="44493-112">**Required**</span></span>|<span data-ttu-id="44493-113">**Description**</span><span class="sxs-lookup"><span data-stu-id="44493-113">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="44493-114">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="44493-114">DefaultValue</span></span>|<span data-ttu-id="44493-115">chaîne (URL)</span><span class="sxs-lookup"><span data-stu-id="44493-115">string (URL)</span></span>|<span data-ttu-id="44493-116">obligatoire</span><span class="sxs-lookup"><span data-stu-id="44493-116">required</span></span>|<span data-ttu-id="44493-117">Spécifie la valeur par défaut de ce paramètre, exprimée pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="44493-117">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="44493-118">Remarques</span><span class="sxs-lookup"><span data-stu-id="44493-118">Remarks</span></span>

<span data-ttu-id="44493-119">Pour un complément de messagerie **, l’icône** > est affichée dans l’interface utilisateur**gérer les compléments** .</span><span class="sxs-lookup"><span data-stu-id="44493-119">For a mail add-in, the icon is displayed in the **File** > **Manage add-ins** UI .</span></span> <span data-ttu-id="44493-120">Pour un complément de contenu ou de volet Office, l’icône s’affiche dans l’interface utilisateur, sous **Insérer** > **Compléments**.</span><span class="sxs-lookup"><span data-stu-id="44493-120">For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span></span>

<span data-ttu-id="44493-121">L’image doit être dans un des formats de fichier suivants : GIF, JPG, PNG, EXIF, BMP ou TIFF.</span><span class="sxs-lookup"><span data-stu-id="44493-121">The image must be in one of the following file formats: GIF, JPG, PNG, EXIF, BMP, or TIFF.</span></span> <span data-ttu-id="44493-122">Pour les applications de contenu et de volet des tâches, la résolution d’image recommandée est de 64 x 64 pixels.</span><span class="sxs-lookup"><span data-stu-id="44493-122">For content and task pane apps, the recommended image resolution is 64 x 64 pixels.</span></span> <span data-ttu-id="44493-123">Pour les applications de messagerie, l’image doit faire 128 x 128 pixels.</span><span class="sxs-lookup"><span data-stu-id="44493-123">For mail apps, the image must be 128 x 128 pixels.</span></span> <span data-ttu-id="44493-124">Pour plus d’informations, voir la section _Créer une identité visuelle cohérente pour votre application_ dans [Création de listings efficaces dans AppSource et dans Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span><span class="sxs-lookup"><span data-stu-id="44493-124">For more information, see the section  _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span></span>
