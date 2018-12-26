---
title: Élément SourceLocation dans le fichier manifeste
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: dc432ebb9482e8e9b8be5d90a838357ccf519ad3
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433515"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="f3e52-102">Élément SourceLocation</span><span class="sxs-lookup"><span data-stu-id="f3e52-102">SourceLocation element</span></span>

<span data-ttu-id="f3e52-p101">Spécifie les emplacements des fichiers source pour votre complément Office sous forme d’URL comprenant entre 1 et 2 018 caractères. L’emplacement source doit être une adresse HTTPS, et non un chemin d’accès de fichier.</span><span class="sxs-lookup"><span data-stu-id="f3e52-p101">Specifies the source file location(s) for your Office Add-in as a URL between 1 and 2018 characters long. The source location must be an HTTPS address, not a file path.</span></span>

<span data-ttu-id="f3e52-105">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="f3e52-105">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="f3e52-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="f3e52-106">Syntax</span></span>

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a><span data-ttu-id="f3e52-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="f3e52-107">Contained in</span></span>

- <span data-ttu-id="f3e52-108">[DefaultSettings](defaultsettings.md) (compléments de contenu et de volet Office)</span><span class="sxs-lookup"><span data-stu-id="f3e52-108">[DefaultSettings](defaultsettings.md) (Content and task pane add-ins)</span></span>
- <span data-ttu-id="f3e52-109">[FormSettings](formsettings.md) (compléments de messagerie)</span><span class="sxs-lookup"><span data-stu-id="f3e52-109">[FormSettings](formsettings.md) (Mail add-ins)</span></span>
- <span data-ttu-id="f3e52-110">[ExtensionPoint](extensionpoint.md) (compléments de messagerie contextuels)</span><span class="sxs-lookup"><span data-stu-id="f3e52-110">[ExtensionPoint](extensionpoint.md) (Contextual mail add-ins)</span></span>

## <a name="can-contain"></a><span data-ttu-id="f3e52-111">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="f3e52-111">Can contain</span></span>

[<span data-ttu-id="f3e52-112">Override</span><span class="sxs-lookup"><span data-stu-id="f3e52-112">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="f3e52-113">Attributs</span><span class="sxs-lookup"><span data-stu-id="f3e52-113">Attributes</span></span>

|<span data-ttu-id="f3e52-114">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="f3e52-114">**Attribute**</span></span>|<span data-ttu-id="f3e52-115">**Type**</span><span class="sxs-lookup"><span data-stu-id="f3e52-115">**Type**</span></span>|<span data-ttu-id="f3e52-116">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="f3e52-116">**Required**</span></span>|<span data-ttu-id="f3e52-117">**Description**</span><span class="sxs-lookup"><span data-stu-id="f3e52-117">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="f3e52-118">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="f3e52-118">DefaultValue</span></span>|<span data-ttu-id="f3e52-119">URL</span><span class="sxs-lookup"><span data-stu-id="f3e52-119">URL</span></span>|<span data-ttu-id="f3e52-120">obligatoire</span><span class="sxs-lookup"><span data-stu-id="f3e52-120">required</span></span>|<span data-ttu-id="f3e52-121">Spécifie la valeur par défaut de ce paramètre pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="f3e52-121">Specifies the default value for this setting for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
