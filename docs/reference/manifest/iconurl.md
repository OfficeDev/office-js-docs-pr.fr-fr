---
title: Élément IconUrl dans le fichier manifeste
description: ''
ms.date: 05/20/2019
localization_priority: Normal
ms.openlocfilehash: 0f518741f0139c9cb240196592edae22b1b09ee7
ms.sourcegitcommit: b0e71ae0ae09c57b843d4de277081845c108a645
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/21/2019
ms.locfileid: "34337201"
---
# <a name="iconurl-element"></a><span data-ttu-id="4463c-102">IconUrl, élément</span><span class="sxs-lookup"><span data-stu-id="4463c-102">IconUrl element</span></span>

<span data-ttu-id="4463c-103">Spécifie l’URL de l’image utilisée pour représenter votre complément Office dans l’UX d’insertion UX et l’Office Store.</span><span class="sxs-lookup"><span data-stu-id="4463c-103">Specifies the URL of the image that is used to represent your Office Add-in in the insertion UX and Office Store.</span></span>

<span data-ttu-id="4463c-104">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="4463c-104">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="4463c-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="4463c-105">Syntax</span></span>

```XML
<IconUrl DefaultValue="string" />
```

## <a name="can-contain"></a><span data-ttu-id="4463c-106">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="4463c-106">Can contain</span></span>

[<span data-ttu-id="4463c-107">Override</span><span class="sxs-lookup"><span data-stu-id="4463c-107">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="4463c-108">Attributs</span><span class="sxs-lookup"><span data-stu-id="4463c-108">Attributes</span></span>

|<span data-ttu-id="4463c-109">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="4463c-109">**Attribute**</span></span>|<span data-ttu-id="4463c-110">**Type**</span><span class="sxs-lookup"><span data-stu-id="4463c-110">**Type**</span></span>|<span data-ttu-id="4463c-111">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="4463c-111">**Required**</span></span>|<span data-ttu-id="4463c-112">**Description**</span><span class="sxs-lookup"><span data-stu-id="4463c-112">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="4463c-113">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="4463c-113">DefaultValue</span></span>|<span data-ttu-id="4463c-114">chaîne</span><span class="sxs-lookup"><span data-stu-id="4463c-114">string</span></span>|<span data-ttu-id="4463c-115">obligatoire</span><span class="sxs-lookup"><span data-stu-id="4463c-115">required</span></span>|<span data-ttu-id="4463c-116">Spécifie la valeur par défaut de ce paramètre, exprimée pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="4463c-116">Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|

## <a name="remarks"></a><span data-ttu-id="4463c-117">Remarques</span><span class="sxs-lookup"><span data-stu-id="4463c-117">Remarks</span></span>

<span data-ttu-id="4463c-p101">Pour un complément de messagerie, l’icône s’affiche dans l’interface utilisateur, sous **Fichier**  >  **Gérer les compléments** (Outlook) ou sous **Paramètres**  >  **Gérer les compléments** (Outlook Web App). Pour un complément de contenu ou de volet Office, l’icône s’affiche dans l’interface utilisateur, sous **Insérer**  >  **Compléments**. Pour tous les types de compléments, l’icône est également utilisée sur le site de l’Office Store si vous publiez votre complément dans l’Office Store.</span><span class="sxs-lookup"><span data-stu-id="4463c-p101">For a mail add-in, the icon is displayed in the  **File** > **Manage add-ins** UI (Outlook) or **Settings** > **Manage add-ins** UI (Outlook Web App). For a content or task pane add-in, the icon is displayed in the **Insert** > **Add-ins** UI. For all add-in types, the icon is also used on the Office Store site, if you publish your add-in to the Office Store.</span></span>

<span data-ttu-id="4463c-121">L’image doit être dans un des formats de fichier suivants : GIF, JPG, PNG, EXIF, BMP ou TIFF.</span><span class="sxs-lookup"><span data-stu-id="4463c-121">The image must be in one of the following file formats: GIF, JPG, PNG, EXIF, BMP, or TIFF.</span></span> <span data-ttu-id="4463c-122">Pour les applications de volet de tâches et de contenu, l’image spécifiée doit contenir 32 x 32 pixels.</span><span class="sxs-lookup"><span data-stu-id="4463c-122">For content and task pane apps, the image specified must be 32 x 32 pixels.</span></span> <span data-ttu-id="4463c-123">Pour les applications de messagerie, la résolution d’image recommandée est de 64 x 64 pixels.</span><span class="sxs-lookup"><span data-stu-id="4463c-123">For mail apps, the recommended image resolution is 64 x 64 pixels.</span></span> <span data-ttu-id="4463c-124">Vous devez également spécifier une icône pour une utilisation avec les applications hôte Office en cours d’exécution sur des écrans haute résolution (DPI) à l’aide de l’élément [HighResolutionIconUrl](highresolutioniconurl.md).</span><span class="sxs-lookup"><span data-stu-id="4463c-124">You should also specify an icon for use with Office host applications running on high DPI screens using the [HighResolutionIconUrl](highresolutioniconurl.md) element.</span></span> <span data-ttu-id="4463c-125">Pour plus d’informations, reportez-vous à la section _Créer une identité visuelle cohérente pour votre application_ dans [Création de listings efficaces dans AppSource et dans Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span><span class="sxs-lookup"><span data-stu-id="4463c-125">For more information, see the section _Create a consistent visual identity for your app_ in [Create effective listings in AppSource and within Office](/office/dev/store/create-effective-office-store-listings#create-a-consistent-visual-identity).</span></span>

<span data-ttu-id="4463c-126">La modification de la valeur `IconUrl` de l’élément au moment de l’exécution n’est actuellement pas prise en charge.</span><span class="sxs-lookup"><span data-stu-id="4463c-126">Changing the value of the `IconUrl` element at runtime is not currently supported.</span></span>