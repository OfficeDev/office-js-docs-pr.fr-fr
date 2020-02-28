---
title: Élément HighResolutionIconUrl dans le fichier manifeste
description: ''
ms.date: 12/04/2018
localization_priority: Normal
ms.openlocfilehash: 41008be6b60d260bef78808af2b8dee1fbd0864a
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325268"
---
# <a name="highresolutioniconurl-element"></a><span data-ttu-id="f740e-102">HighResolutionIconUrl, élément</span><span class="sxs-lookup"><span data-stu-id="f740e-102">HighResolutionIconUrl element</span></span>

<span data-ttu-id="f740e-103">Spécifie l’URL de l’image qui est utilisée pour représenter votre complément Office dans l’UX d’insertion UX et l’Office Store sur les écrans à haute résolution (DPI).</span><span class="sxs-lookup"><span data-stu-id="f740e-103">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store on high DPI screens.</span></span>

<span data-ttu-id="f740e-104">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="f740e-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="f740e-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="f740e-105">Syntax</span></span>

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="f740e-106">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="f740e-106">Can contain</span></span>

[<span data-ttu-id="f740e-107">Override</span><span class="sxs-lookup"><span data-stu-id="f740e-107">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="f740e-108">Attributs</span><span class="sxs-lookup"><span data-stu-id="f740e-108">Attributes</span></span>

|<span data-ttu-id="f740e-109">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="f740e-109">**Attribute**</span></span>|<span data-ttu-id="f740e-110">**Type**</span><span class="sxs-lookup"><span data-stu-id="f740e-110">**Type**</span></span>|<span data-ttu-id="f740e-111">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="f740e-111">**Required**</span></span>|<span data-ttu-id="f740e-112">**Description**</span><span class="sxs-lookup"><span data-stu-id="f740e-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="f740e-113">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="f740e-113">DefaultValue</span></span>|<span data-ttu-id="f740e-114">chaîne (URL)</span><span class="sxs-lookup"><span data-stu-id="f740e-114">string (URL)</span></span>|<span data-ttu-id="f740e-115">obligatoire</span><span class="sxs-lookup"><span data-stu-id="f740e-115">required</span></span>|<span data-ttu-id="f740e-116">Spécifie la valeur par défaut de ce paramètre, exprimée pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="f740e-116">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="f740e-117">Remarques</span><span class="sxs-lookup"><span data-stu-id="f740e-117">Remarks</span></span>

<span data-ttu-id="f740e-118">Pour un complément de messagerie **, l’icône** > est affichée dans l’interface utilisateur**gérer les compléments** .</span><span class="sxs-lookup"><span data-stu-id="f740e-118">For a mail add-in, the icon is displayed in the **File** > **Manage add-ins** UI .</span></span> <span data-ttu-id="f740e-119">Pour un complément de contenu ou de volet Office, l’icône s’affiche dans l’interface utilisateur, sous **Insérer** > **Compléments**.</span><span class="sxs-lookup"><span data-stu-id="f740e-119">For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span></span>

<span data-ttu-id="f740e-120">L’image doit être dans un des formats de fichier suivants : GIF, JPG, PNG, EXIF, BMP ou TIFF.</span><span class="sxs-lookup"><span data-stu-id="f740e-120">The image must be in one of the following file formats: GIF, JPG, PNG, EXIF, BMP, or TIFF.</span></span> <span data-ttu-id="f740e-121">Pour les applications de contenu et de volet des tâches, la résolution d’image recommandée est de 64 x 64 pixels.</span><span class="sxs-lookup"><span data-stu-id="f740e-121">For content and task pane apps, the recommended image resolution is 64 x 64 pixels.</span></span> <span data-ttu-id="f740e-122">Pour les applications de messagerie, l’image doit faire 128 x 128 pixels.</span><span class="sxs-lookup"><span data-stu-id="f740e-122">For mail apps, the image must be 128 x 128 pixels.</span></span> <span data-ttu-id="f740e-123">Pour plus d’informations, voir la section _Créer une identité visuelle cohérente pour votre application_ dans [Création de listings efficaces dans AppSource et dans Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span><span class="sxs-lookup"><span data-stu-id="f740e-123">For more information, see the section  _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span></span>
