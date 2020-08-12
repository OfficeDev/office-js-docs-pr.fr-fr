---
title: Élément SourceLocation dans le fichier manifeste
description: L’élément SourceLocation spécifie les emplacements des fichiers source pour votre complément Office.
ms.date: 05/12/2020
localization_priority: Normal
ms.openlocfilehash: 447adb7df7d0c59305fe5046357959fcd7824735
ms.sourcegitcommit: cc6886b47c84ac37a3c957ff85dd0ed526ca5e43
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/12/2020
ms.locfileid: "46641402"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="ba240-103">Élément SourceLocation</span><span class="sxs-lookup"><span data-stu-id="ba240-103">SourceLocation element</span></span>

<span data-ttu-id="ba240-104">Spécifie les emplacements des fichiers source pour votre complément Office sous la forme d’une URL de 1 à 2018 caractères.</span><span class="sxs-lookup"><span data-stu-id="ba240-104">Specifies the source file locations for your Office Add-in as a URL between 1 and 2018 characters long.</span></span> <span data-ttu-id="ba240-105">L’emplacement source doit être une adresse HTTPS, et non un chemin d’accès de fichier.</span><span class="sxs-lookup"><span data-stu-id="ba240-105">The source location must be an HTTPS address, not a file path.</span></span>

<span data-ttu-id="ba240-106">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="ba240-106">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="ba240-107">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="ba240-107">Syntax</span></span>

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a><span data-ttu-id="ba240-108">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="ba240-108">Contained in</span></span>

- <span data-ttu-id="ba240-109">[DefaultSettings](defaultsettings.md) (compléments de contenu et de volet Office)</span><span class="sxs-lookup"><span data-stu-id="ba240-109">[DefaultSettings](defaultsettings.md) (Content and task pane add-ins)</span></span>
- <span data-ttu-id="ba240-110">[FormSettings](formsettings.md) (compléments de messagerie)</span><span class="sxs-lookup"><span data-stu-id="ba240-110">[FormSettings](formsettings.md) (Mail add-ins)</span></span>
- <span data-ttu-id="ba240-111">[ExtensionPoint](extensionpoint.md) (contextuel et LaunchEvent (aperçu) des compléments de messagerie)</span><span class="sxs-lookup"><span data-stu-id="ba240-111">[ExtensionPoint](extensionpoint.md) (Contextual and LaunchEvent (preview) mail add-ins)</span></span>

## <a name="can-contain"></a><span data-ttu-id="ba240-112">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="ba240-112">Can contain</span></span>

[<span data-ttu-id="ba240-113">Override</span><span class="sxs-lookup"><span data-stu-id="ba240-113">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="ba240-114">Attributs</span><span class="sxs-lookup"><span data-stu-id="ba240-114">Attributes</span></span>

|<span data-ttu-id="ba240-115">Attribut</span><span class="sxs-lookup"><span data-stu-id="ba240-115">Attribute</span></span>|<span data-ttu-id="ba240-116">Type</span><span class="sxs-lookup"><span data-stu-id="ba240-116">Type</span></span>|<span data-ttu-id="ba240-117">Requis</span><span class="sxs-lookup"><span data-stu-id="ba240-117">Required</span></span>|<span data-ttu-id="ba240-118">Description</span><span class="sxs-lookup"><span data-stu-id="ba240-118">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="ba240-119">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="ba240-119">DefaultValue</span></span>|<span data-ttu-id="ba240-120">URL</span><span class="sxs-lookup"><span data-stu-id="ba240-120">URL</span></span>|<span data-ttu-id="ba240-121">obligatoire</span><span class="sxs-lookup"><span data-stu-id="ba240-121">required</span></span>|<span data-ttu-id="ba240-122">Spécifie la valeur par défaut de ce paramètre pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="ba240-122">Specifies the default value for this setting for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
