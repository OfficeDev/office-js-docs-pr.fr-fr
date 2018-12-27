---
title: Élément HighResolutionIconUrl dans le fichier manifeste
description: ''
ms.date: 12/04/2018
ms.openlocfilehash: dc8feb92eb8a53351679834a39c012b47f43aad4
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432591"
---
# <a name="highresolutioniconurl-element"></a><span data-ttu-id="f8649-102">HighResolutionIconUrl, élément</span><span class="sxs-lookup"><span data-stu-id="f8649-102">HighResolutionIconUrl element</span></span>

<span data-ttu-id="f8649-103">Spécifie l’URL de l’image qui est utilisée pour représenter votre complément Office dans l’UX d’insertion UX et l’Office Store sur les écrans à haute résolution (DPI).</span><span class="sxs-lookup"><span data-stu-id="f8649-103">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store on high DPI screens.</span></span>

<span data-ttu-id="f8649-104">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="f8649-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="f8649-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="f8649-105">Syntax</span></span>

```XML
<HighResolutionIconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="f8649-106">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="f8649-106">Can contain</span></span>

[<span data-ttu-id="f8649-107">Override</span><span class="sxs-lookup"><span data-stu-id="f8649-107">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="f8649-108">Attributs</span><span class="sxs-lookup"><span data-stu-id="f8649-108">Attributes</span></span>

|<span data-ttu-id="f8649-109">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="f8649-109">**Attribute**</span></span>|<span data-ttu-id="f8649-110">**Type**</span><span class="sxs-lookup"><span data-stu-id="f8649-110">**Type**</span></span>|<span data-ttu-id="f8649-111">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="f8649-111">**Required**</span></span>|<span data-ttu-id="f8649-112">**Description**</span><span class="sxs-lookup"><span data-stu-id="f8649-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="f8649-113">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="f8649-113">DefaultValue</span></span>|<span data-ttu-id="f8649-114">chaîne (URL)</span><span class="sxs-lookup"><span data-stu-id="f8649-114">string (URL)</span></span>|<span data-ttu-id="f8649-115">obligatoire</span><span class="sxs-lookup"><span data-stu-id="f8649-115">required</span></span>|<span data-ttu-id="f8649-116">Spécifie la valeur par défaut de ce paramètre, exprimée pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="f8649-116">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="f8649-117">Remarques</span><span class="sxs-lookup"><span data-stu-id="f8649-117">Remarks</span></span>

<span data-ttu-id="f8649-p101">Pour un complément de messagerie, l’icône apparaît dans l’interface utilisateur, sous **Fichier**  >  **Gérer les compléments**. Pour un complément de contenu ou de volet Office, l’icône apparaît dans l’interface utilisateur, sous **Insérer**  >  **Compléments**.</span><span class="sxs-lookup"><span data-stu-id="f8649-p101">For a mail add-in, the icon is displayed in the  **File** > **Manage add-ins** UI . For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span></span>

<span data-ttu-id="f8649-120">L’image doit être dans un des formats de fichier suivants : GIF, JPG, PNG, EXIF, BMP ou TIFF.</span><span class="sxs-lookup"><span data-stu-id="f8649-120">The image must be in one of the following file formats at a recommended resolution of 64 x 64 pixels: GIF, JPG, PNG, EXIF, BMP or TIFF.</span></span> <span data-ttu-id="f8649-121">Pour les applications de contenu et de volet des tâches, la résolution d’image recommandée est de 64 x 64 pixels.</span><span class="sxs-lookup"><span data-stu-id="f8649-121">For content and task pane apps, the recommended image resolution is 64 x 64 pixels.</span></span> <span data-ttu-id="f8649-122">Pour les applications de messagerie, l’image doit faire 128 x 128 pixels.</span><span class="sxs-lookup"><span data-stu-id="f8649-122">For mail apps, the image must be 128 x 128 pixels.</span></span> <span data-ttu-id="f8649-123">Pour plus d’informations, voir la section _Créer une identité visuelle cohérente pour votre application_ dans [Création de listings efficaces dans AppSource et dans Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span><span class="sxs-lookup"><span data-stu-id="f8649-123">For more information, see the section  Create a consistent visual identity for your app in Create effective Office Store apps and add-ins.</span></span>
