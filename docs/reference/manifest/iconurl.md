---
title: Élément IconUrl dans le fichier manifeste
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 44992a3c5f9ceba55b09f4b14e36b5b2935ee669
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771799"
---
# <a name="iconurl-element"></a><span data-ttu-id="eee3f-102">IconUrl, élément</span><span class="sxs-lookup"><span data-stu-id="eee3f-102">IconUrl element</span></span>

<span data-ttu-id="eee3f-103">Spécifie l’URL de l’image utilisée pour représenter votre complément Office dans l’UX d’insertion UX et l’Office Store.</span><span class="sxs-lookup"><span data-stu-id="eee3f-103">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store.</span></span>

<span data-ttu-id="eee3f-104">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="eee3f-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="eee3f-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="eee3f-105">Syntax</span></span>

```XML
<IconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="eee3f-106">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="eee3f-106">Can contain</span></span>

[<span data-ttu-id="eee3f-107">Override</span><span class="sxs-lookup"><span data-stu-id="eee3f-107">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="eee3f-108">Attributs</span><span class="sxs-lookup"><span data-stu-id="eee3f-108">Attributes</span></span>

|<span data-ttu-id="eee3f-109">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="eee3f-109">**Attribute**</span></span>|<span data-ttu-id="eee3f-110">**Type**</span><span class="sxs-lookup"><span data-stu-id="eee3f-110">**Type**</span></span>|<span data-ttu-id="eee3f-111">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="eee3f-111">**Required**</span></span>|<span data-ttu-id="eee3f-112">**Description**</span><span class="sxs-lookup"><span data-stu-id="eee3f-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="eee3f-113">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="eee3f-113">DefaultValue</span></span>|<span data-ttu-id="eee3f-114">chaîne</span><span class="sxs-lookup"><span data-stu-id="eee3f-114">string</span></span>|<span data-ttu-id="eee3f-115">obligatoire</span><span class="sxs-lookup"><span data-stu-id="eee3f-115">required</span></span>|<span data-ttu-id="eee3f-116">Spécifie la valeur par défaut de ce paramètre, exprimée pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="eee3f-116">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="eee3f-117">Remarques</span><span class="sxs-lookup"><span data-stu-id="eee3f-117">Remarks</span></span>

<span data-ttu-id="eee3f-118">Pour un complément de messagerie, l’icône est affichée dans \*\*\*\* > l’interface utilisateur de**gestion** des compléments (Outlook) ou **paramètres** > **gérer les compléments** (Outlook sur le Web).</span><span class="sxs-lookup"><span data-stu-id="eee3f-118">For a mail add-in, the icon is displayed in the  **File** > **Manage add-ins** UI (Outlook) or **Settings** > **Manage add-ins** UI (Outlook on the web).</span></span> <span data-ttu-id="eee3f-119">Pour un complément de contenu ou de volet Office, l’icône s’affiche dans l’interface utilisateur, sous **Insérer** > **Compléments**.</span><span class="sxs-lookup"><span data-stu-id="eee3f-119">For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI.</span></span> <span data-ttu-id="eee3f-120">Pour tous les types de complément, l’icône est également utilisée dans [AppSource](https://appsource.microsoft.com), si vous publiez votre complément dans AppSource.</span><span class="sxs-lookup"><span data-stu-id="eee3f-120">For all add-in types, the icon is also used in [AppSource](https://appsource.microsoft.com), if you publish your add-in to AppSource.</span></span>

<span data-ttu-id="eee3f-121">L’image doit être dans un des formats de fichier suivants : GIF, JPG, PNG, EXIF, BMP ou TIFF.</span><span class="sxs-lookup"><span data-stu-id="eee3f-121">The image must be in one of the following file formats: GIF, JPG, PNG, EXIF, BMP, or TIFF.</span></span> <span data-ttu-id="eee3f-122">Pour les applications de volet de tâches et de contenu, l’image spécifiée doit contenir 32 x 32 pixels.</span><span class="sxs-lookup"><span data-stu-id="eee3f-122">For content and task pane apps, the image specified must be 32 x 32 pixels.</span></span> <span data-ttu-id="eee3f-123">Pour les applications de messagerie, la résolution d’image recommandée est de 64 x 64 pixels.</span><span class="sxs-lookup"><span data-stu-id="eee3f-123">For mail apps, the recommended image resolution is 64 x 64 pixels.</span></span> <span data-ttu-id="eee3f-124">Vous devez également spécifier une icône pour une utilisation avec les applications hôte Office en cours d’exécution sur des écrans haute résolution (DPI) à l’aide de l’élément [HighResolutionIconUrl](highresolutioniconurl.md).</span><span class="sxs-lookup"><span data-stu-id="eee3f-124">You should also specify an icon for use with Office host applications running on high DPI screens using the [HighResolutionIconUrl](highresolutioniconurl.md) element.</span></span> <span data-ttu-id="eee3f-125">Pour plus d’informations, reportez-vous à la section _Créer une identité visuelle cohérente pour votre application_ dans [Création de listings efficaces dans AppSource et dans Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span><span class="sxs-lookup"><span data-stu-id="eee3f-125">For more information, see the section _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span></span>

<span data-ttu-id="eee3f-126">La modification de la valeur `IconUrl` de l’élément au moment de l’exécution n’est actuellement pas prise en charge.</span><span class="sxs-lookup"><span data-stu-id="eee3f-126">Changing the value of the `IconUrl` element at runtime is not currently supported.</span></span>