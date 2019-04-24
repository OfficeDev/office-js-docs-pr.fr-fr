---
title: Élément HighResolutionIconUrl dans le fichier manifeste
description: ''
ms.date: 12/04/2018
localization_priority: Normal
ms.openlocfilehash: 5264fc969bda30a9b2212996800b984533a3188c
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32452087"
---
# <a name="highresolutioniconurl-element"></a><span data-ttu-id="19902-102">HighResolutionIconUrl, élément</span><span class="sxs-lookup"><span data-stu-id="19902-102">HighResolutionIconUrl element</span></span>

<span data-ttu-id="19902-103">Spécifie l’URL de l’image qui est utilisée pour représenter votre complément Office dans l’UX d’insertion UX et l’Office Store sur les écrans à haute résolution (DPI).</span><span class="sxs-lookup"><span data-stu-id="19902-103">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store on high DPI screens.</span></span>

<span data-ttu-id="19902-104">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="19902-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="19902-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="19902-105">Syntax</span></span>

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="19902-106">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="19902-106">Can contain</span></span>

[<span data-ttu-id="19902-107">Override</span><span class="sxs-lookup"><span data-stu-id="19902-107">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="19902-108">Attributs</span><span class="sxs-lookup"><span data-stu-id="19902-108">Attributes</span></span>

|<span data-ttu-id="19902-109">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="19902-109">**Attribute**</span></span>|<span data-ttu-id="19902-110">**Type**</span><span class="sxs-lookup"><span data-stu-id="19902-110">**Type**</span></span>|<span data-ttu-id="19902-111">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="19902-111">**Required**</span></span>|<span data-ttu-id="19902-112">**Description**</span><span class="sxs-lookup"><span data-stu-id="19902-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="19902-113">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="19902-113">DefaultValue</span></span>|<span data-ttu-id="19902-114">chaîne (URL)</span><span class="sxs-lookup"><span data-stu-id="19902-114">string (URL)</span></span>|<span data-ttu-id="19902-115">obligatoire</span><span class="sxs-lookup"><span data-stu-id="19902-115">required</span></span>|<span data-ttu-id="19902-116">Spécifie la valeur par défaut de ce paramètre, exprimée pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="19902-116">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="19902-117">Remarques</span><span class="sxs-lookup"><span data-stu-id="19902-117">Remarks</span></span>

<span data-ttu-id="19902-p101">Pour un complément de messagerie, l’icône apparaît dans l’interface utilisateur, sous **Fichier**  >  **Gérer les compléments**. Pour un complément de contenu ou de volet Office, l’icône apparaît dans l’interface utilisateur, sous **Insérer**  >  **Compléments**.</span><span class="sxs-lookup"><span data-stu-id="19902-p101">For a mail add-in, the icon is displayed in the  **File** > **Manage add-ins** UI . For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span></span>

<span data-ttu-id="19902-120">L’image doit être dans un des formats de fichier suivants : GIF, JPG, PNG, EXIF, BMP ou TIFF.</span><span class="sxs-lookup"><span data-stu-id="19902-120">The image must be in one of the following file formats: GIF, JPG, PNG, EXIF, BMP, or TIFF.</span></span> <span data-ttu-id="19902-121">Pour les applications de contenu et de volet des tâches, la résolution d’image recommandée est de 64 x 64 pixels.</span><span class="sxs-lookup"><span data-stu-id="19902-121">For content and task pane apps, the recommended image resolution is 64 x 64 pixels.</span></span> <span data-ttu-id="19902-122">Pour les applications de messagerie, l’image doit faire 128 x 128 pixels.</span><span class="sxs-lookup"><span data-stu-id="19902-122">For mail apps, the image must be 128 x 128 pixels.</span></span> <span data-ttu-id="19902-123">Pour plus d’informations, voir la section _Créer une identité visuelle cohérente pour votre application_ dans [Création de listings efficaces dans AppSource et dans Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span><span class="sxs-lookup"><span data-stu-id="19902-123">For more information, see the section  _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span></span>
