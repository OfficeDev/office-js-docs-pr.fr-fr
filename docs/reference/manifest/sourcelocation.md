---
title: Élément SourceLocation dans le fichier manifeste
description: L’élément SourceLocation spécifie les emplacements de fichiers sources pour votre Office de recherche.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 4dcd093db2f23220eaa34c0c81300c4994c1a697
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590896"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="7b2ad-103">Élément SourceLocation</span><span class="sxs-lookup"><span data-stu-id="7b2ad-103">SourceLocation element</span></span>

<span data-ttu-id="7b2ad-104">Spécifie les emplacements de fichiers sources de votre Office sous la mesure d’une URL de 1 à 2 018 caractères.</span><span class="sxs-lookup"><span data-stu-id="7b2ad-104">Specifies the source file locations for your Office Add-in as a URL between 1 and 2018 characters long.</span></span> <span data-ttu-id="7b2ad-105">L’emplacement source doit être une adresse HTTPS, et non un chemin d’accès de fichier.</span><span class="sxs-lookup"><span data-stu-id="7b2ad-105">The source location must be an HTTPS address, not a file path.</span></span>

<span data-ttu-id="7b2ad-106">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="7b2ad-106">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="7b2ad-107">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="7b2ad-107">Syntax</span></span>

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a><span data-ttu-id="7b2ad-108">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="7b2ad-108">Contained in</span></span>

- <span data-ttu-id="7b2ad-109">[DefaultSettings](defaultsettings.md) (compléments de contenu et de volet Office)</span><span class="sxs-lookup"><span data-stu-id="7b2ad-109">[DefaultSettings](defaultsettings.md) (Content and task pane add-ins)</span></span>
- <span data-ttu-id="7b2ad-110">[FormSettings](formsettings.md) (compléments de messagerie)</span><span class="sxs-lookup"><span data-stu-id="7b2ad-110">[FormSettings](formsettings.md) (Mail add-ins)</span></span>
- <span data-ttu-id="7b2ad-111">[ExtensionPoint](extensionpoint.md) (modules de messagerie contextuels et LaunchEvent)</span><span class="sxs-lookup"><span data-stu-id="7b2ad-111">[ExtensionPoint](extensionpoint.md) (Contextual and LaunchEvent mail add-ins)</span></span>

## <a name="can-contain"></a><span data-ttu-id="7b2ad-112">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="7b2ad-112">Can contain</span></span>

[<span data-ttu-id="7b2ad-113">Override</span><span class="sxs-lookup"><span data-stu-id="7b2ad-113">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="7b2ad-114">Attributs</span><span class="sxs-lookup"><span data-stu-id="7b2ad-114">Attributes</span></span>

|<span data-ttu-id="7b2ad-115">Attribut</span><span class="sxs-lookup"><span data-stu-id="7b2ad-115">Attribute</span></span>|<span data-ttu-id="7b2ad-116">Type</span><span class="sxs-lookup"><span data-stu-id="7b2ad-116">Type</span></span>|<span data-ttu-id="7b2ad-117">Requis</span><span class="sxs-lookup"><span data-stu-id="7b2ad-117">Required</span></span>|<span data-ttu-id="7b2ad-118">Description</span><span class="sxs-lookup"><span data-stu-id="7b2ad-118">Description</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="7b2ad-119">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="7b2ad-119">DefaultValue</span></span>|<span data-ttu-id="7b2ad-120">URL</span><span class="sxs-lookup"><span data-stu-id="7b2ad-120">URL</span></span>|<span data-ttu-id="7b2ad-121">obligatoire</span><span class="sxs-lookup"><span data-stu-id="7b2ad-121">required</span></span>|<span data-ttu-id="7b2ad-122">Spécifie la valeur par défaut de ce paramètre pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="7b2ad-122">Specifies the default value for this setting for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
