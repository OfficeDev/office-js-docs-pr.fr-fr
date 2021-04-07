---
title: Élément IconUrl dans le fichier manifeste
description: L’élément IconUrl spécifie l’URL de l’image qui représente votre add-in Office dans l’UX d’insertion et l’Office Store.
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: 68a449b40f6084d26140d59fec61967e163196df
ms.sourcegitcommit: 0bff0411d8cfefd4bb00c189643358e6fb1df95e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/07/2021
ms.locfileid: "51604638"
---
# <a name="iconurl-element"></a><span data-ttu-id="bf51a-103">IconUrl, élément</span><span class="sxs-lookup"><span data-stu-id="bf51a-103">IconUrl element</span></span>

<span data-ttu-id="bf51a-104">Spécifie l’URL de l’image utilisée pour représenter votre complément Office dans l’UX d’insertion UX et l’Office Store.</span><span class="sxs-lookup"><span data-stu-id="bf51a-104">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store.</span></span>

<span data-ttu-id="bf51a-105">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="bf51a-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="bf51a-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="bf51a-106">Syntax</span></span>

```XML
<IconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="bf51a-107">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="bf51a-107">Can contain</span></span>

[<span data-ttu-id="bf51a-108">Override</span><span class="sxs-lookup"><span data-stu-id="bf51a-108">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="bf51a-109">Attributs</span><span class="sxs-lookup"><span data-stu-id="bf51a-109">Attributes</span></span>

|<span data-ttu-id="bf51a-110">Attribut</span><span class="sxs-lookup"><span data-stu-id="bf51a-110">Attribute</span></span>|<span data-ttu-id="bf51a-111">Type</span><span class="sxs-lookup"><span data-stu-id="bf51a-111">Type</span></span>|<span data-ttu-id="bf51a-112">Requis</span><span class="sxs-lookup"><span data-stu-id="bf51a-112">Required</span></span>|<span data-ttu-id="bf51a-113">Description</span><span class="sxs-lookup"><span data-stu-id="bf51a-113">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="bf51a-114">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="bf51a-114">DefaultValue</span></span>|<span data-ttu-id="bf51a-115">chaîne</span><span class="sxs-lookup"><span data-stu-id="bf51a-115">string</span></span>|<span data-ttu-id="bf51a-116">obligatoire</span><span class="sxs-lookup"><span data-stu-id="bf51a-116">required</span></span>|<span data-ttu-id="bf51a-117">Spécifie la valeur par défaut de ce paramètre, exprimée pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="bf51a-117">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="bf51a-118">Remarques</span><span class="sxs-lookup"><span data-stu-id="bf51a-118">Remarks</span></span>

<span data-ttu-id="bf51a-119">Pour un module de messagerie, l’icône s’affiche dans l’interface utilisateur de gestion des fichiers (Outlook) ou dans l’interface utilisateur de gestion des paramètres des applications  >     >   (Outlook sur le web).</span><span class="sxs-lookup"><span data-stu-id="bf51a-119">For a mail add-in, the icon is displayed in the **File** > **Manage add-ins** UI (Outlook) or **Settings** > **Manage add-ins** UI (Outlook on the web).</span></span> <span data-ttu-id="bf51a-120">Pour un complément de contenu ou de volet Office, l’icône s’affiche dans l’interface utilisateur, sous **Insérer** > **Compléments**.</span><span class="sxs-lookup"><span data-stu-id="bf51a-120">For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span></span> <span data-ttu-id="bf51a-121">Pour tous les types de modules, l’icône est également utilisée dans [AppSource,](https://appsource.microsoft.com)si vous publiez votre application dans AppSource.</span><span class="sxs-lookup"><span data-stu-id="bf51a-121">For all add-in types, the icon is also used in [AppSource](https://appsource.microsoft.com), if you publish your add-in to AppSource.</span></span>

<span data-ttu-id="bf51a-122">L’image doit être dans un des formats de fichier suivants : GIF, JPG, PNG, EXIF, BMP ou TIFF.</span><span class="sxs-lookup"><span data-stu-id="bf51a-122">The image must be in one of the following file formats: GIF, JPG, PNG, EXIF, BMP, or TIFF.</span></span> <span data-ttu-id="bf51a-123">Pour les applications de volet de tâches et de contenu, l’image spécifiée doit contenir 32 x 32 pixels.</span><span class="sxs-lookup"><span data-stu-id="bf51a-123">For content and task pane apps, the image specified must be 32 x 32 pixels.</span></span> <span data-ttu-id="bf51a-124">Pour les applications de messagerie, la résolution d’image doit être de 64 x 64 pixels.</span><span class="sxs-lookup"><span data-stu-id="bf51a-124">For mail apps, the image resolution must be 64 x 64 pixels.</span></span> <span data-ttu-id="bf51a-125">Vous devez également spécifier une icône à utiliser avec les applications clientes Office qui s’exécutent sur des écrans HAUTESPI à l’aide de l’élément [HighResolutionIconUrl.](highresolutioniconurl.md)</span><span class="sxs-lookup"><span data-stu-id="bf51a-125">You should also specify an icon for use with Office client applications running on high DPI screens using the [HighResolutionIconUrl](highresolutioniconurl.md) element.</span></span> <span data-ttu-id="bf51a-126">Pour plus d’informations, reportez-vous à la section _Créer une identité visuelle cohérente pour votre application_ dans [Création de listings efficaces dans AppSource et dans Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span><span class="sxs-lookup"><span data-stu-id="bf51a-126">For more information, see the section _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span></span>

<span data-ttu-id="bf51a-127">La modification de la valeur de `IconUrl` l’élément au moment de l’runtime n’est pas prise en charge actuellement.</span><span class="sxs-lookup"><span data-stu-id="bf51a-127">Changing the value of the `IconUrl` element at runtime is not currently supported.</span></span>
