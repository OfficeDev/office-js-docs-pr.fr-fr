---
title: Élément SourceLocation dans le fichier manifeste
description: L’élément SourceLocation spécifie les emplacements des fichiers source pour votre complément Office.
ms.date: 05/12/2020
localization_priority: Normal
ms.openlocfilehash: 9af2337263314bec5ce04eb0d22626ab368c19ef
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608725"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="3e01d-103">Élément SourceLocation</span><span class="sxs-lookup"><span data-stu-id="3e01d-103">SourceLocation element</span></span>

<span data-ttu-id="3e01d-104">Spécifie les emplacements des fichiers source pour votre complément Office sous la forme d’une URL de 1 à 2018 caractères.</span><span class="sxs-lookup"><span data-stu-id="3e01d-104">Specifies the source file locations for your Office Add-in as a URL between 1 and 2018 characters long.</span></span> <span data-ttu-id="3e01d-105">L’emplacement source doit être une adresse HTTPS, et non un chemin d’accès de fichier.</span><span class="sxs-lookup"><span data-stu-id="3e01d-105">The source location must be an HTTPS address, not a file path.</span></span>

<span data-ttu-id="3e01d-106">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="3e01d-106">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="3e01d-107">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="3e01d-107">Syntax</span></span>

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a><span data-ttu-id="3e01d-108">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="3e01d-108">Contained in</span></span>

- <span data-ttu-id="3e01d-109">[DefaultSettings](defaultsettings.md) (compléments de contenu et de volet Office)</span><span class="sxs-lookup"><span data-stu-id="3e01d-109">[DefaultSettings](defaultsettings.md) (Content and task pane add-ins)</span></span>
- <span data-ttu-id="3e01d-110">[FormSettings](formsettings.md) (compléments de messagerie)</span><span class="sxs-lookup"><span data-stu-id="3e01d-110">[FormSettings](formsettings.md) (Mail add-ins)</span></span>
- <span data-ttu-id="3e01d-111">[ExtensionPoint](extensionpoint.md) (contextuel et LaunchEvent (aperçu) des compléments de messagerie)</span><span class="sxs-lookup"><span data-stu-id="3e01d-111">[ExtensionPoint](extensionpoint.md) (Contextual and LaunchEvent (preview) mail add-ins)</span></span>

## <a name="can-contain"></a><span data-ttu-id="3e01d-112">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="3e01d-112">Can contain</span></span>

[<span data-ttu-id="3e01d-113">Override</span><span class="sxs-lookup"><span data-stu-id="3e01d-113">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="3e01d-114">Attributs</span><span class="sxs-lookup"><span data-stu-id="3e01d-114">Attributes</span></span>

|<span data-ttu-id="3e01d-115">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="3e01d-115">**Attribute**</span></span>|<span data-ttu-id="3e01d-116">**Type**</span><span class="sxs-lookup"><span data-stu-id="3e01d-116">**Type**</span></span>|<span data-ttu-id="3e01d-117">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="3e01d-117">**Required**</span></span>|<span data-ttu-id="3e01d-118">**Description**</span><span class="sxs-lookup"><span data-stu-id="3e01d-118">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="3e01d-119">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="3e01d-119">DefaultValue</span></span>|<span data-ttu-id="3e01d-120">URL</span><span class="sxs-lookup"><span data-stu-id="3e01d-120">URL</span></span>|<span data-ttu-id="3e01d-121">obligatoire</span><span class="sxs-lookup"><span data-stu-id="3e01d-121">required</span></span>|<span data-ttu-id="3e01d-122">Spécifie la valeur par défaut de ce paramètre pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="3e01d-122">Specifies the default value for this setting for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
