---
title: Élément SourceLocation dans le fichier manifeste
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 7544e2bae480b9431c8912533ea1b761132a355e
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451975"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="46a6a-102">Élément SourceLocation</span><span class="sxs-lookup"><span data-stu-id="46a6a-102">SourceLocation element</span></span>

<span data-ttu-id="46a6a-p101">Spécifie les emplacements des fichiers source pour votre complément Office sous forme d’URL comprenant entre 1 et 2 018 caractères. L’emplacement source doit être une adresse HTTPS, et non un chemin d’accès de fichier.</span><span class="sxs-lookup"><span data-stu-id="46a6a-p101">Specifies the source file location(s) for your Office Add-in as a URL between 1 and 2018 characters long. The source location must be an HTTPS address, not a file path.</span></span>

<span data-ttu-id="46a6a-105">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="46a6a-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="46a6a-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="46a6a-106">Syntax</span></span>

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a><span data-ttu-id="46a6a-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="46a6a-107">Contained in</span></span>

- <span data-ttu-id="46a6a-108">[DefaultSettings](defaultsettings.md) (compléments de contenu et de volet Office)</span><span class="sxs-lookup"><span data-stu-id="46a6a-108">[DefaultSettings](defaultsettings.md) (Content and task pane add-ins)</span></span>
- <span data-ttu-id="46a6a-109">[FormSettings](formsettings.md) (compléments de messagerie)</span><span class="sxs-lookup"><span data-stu-id="46a6a-109">[FormSettings](formsettings.md) (Mail add-ins)</span></span>
- <span data-ttu-id="46a6a-110">[ExtensionPoint](extensionpoint.md) (compléments de messagerie contextuels)</span><span class="sxs-lookup"><span data-stu-id="46a6a-110">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins)</span></span>

## <a name="can-contain"></a><span data-ttu-id="46a6a-111">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="46a6a-111">Can contain</span></span>

[<span data-ttu-id="46a6a-112">Override</span><span class="sxs-lookup"><span data-stu-id="46a6a-112">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="46a6a-113">Attributs</span><span class="sxs-lookup"><span data-stu-id="46a6a-113">Attributes</span></span>

|<span data-ttu-id="46a6a-114">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="46a6a-114">**Attribute**</span></span>|<span data-ttu-id="46a6a-115">**Type**</span><span class="sxs-lookup"><span data-stu-id="46a6a-115">**Type**</span></span>|<span data-ttu-id="46a6a-116">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="46a6a-116">**Required**</span></span>|<span data-ttu-id="46a6a-117">**Description**</span><span class="sxs-lookup"><span data-stu-id="46a6a-117">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="46a6a-118">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="46a6a-118">DefaultValue</span></span>|<span data-ttu-id="46a6a-119">URL</span><span class="sxs-lookup"><span data-stu-id="46a6a-119">URL</span></span>|<span data-ttu-id="46a6a-120">obligatoire</span><span class="sxs-lookup"><span data-stu-id="46a6a-120">required</span></span>|<span data-ttu-id="46a6a-121">Spécifie la valeur par défaut de ce paramètre pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="46a6a-121">Specifies the default value for this setting for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
