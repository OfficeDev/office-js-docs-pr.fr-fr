---
title: Élément SourceLocation dans le fichier manifeste
description: L’élément SourceLocation spécifie les emplacements des fichiers source pour votre complément Office.
ms.date: 05/12/2020
localization_priority: Normal
ms.openlocfilehash: 642780c3231523ea579ca548b3f3f984b2856666
ms.sourcegitcommit: f62d9630de69c5c070e3d4048205f5cc654db7e4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/18/2020
ms.locfileid: "44278398"
---
# <a name="sourcelocation-element"></a><span data-ttu-id="7b451-103">Élément SourceLocation</span><span class="sxs-lookup"><span data-stu-id="7b451-103">SourceLocation element</span></span>

<span data-ttu-id="7b451-104">Spécifie les emplacements des fichiers source pour votre complément Office sous la forme d’une URL de 1 à 2018 caractères.</span><span class="sxs-lookup"><span data-stu-id="7b451-104">Specifies the source file locations for your Office Add-in as a URL between 1 and 2018 characters long.</span></span> <span data-ttu-id="7b451-105">L’emplacement source doit être une adresse HTTPS, et non un chemin d’accès de fichier.</span><span class="sxs-lookup"><span data-stu-id="7b451-105">The source location must be an HTTPS address, not a file path.</span></span>

<span data-ttu-id="7b451-106">**Type de complément :** application de contenu, de volet Office, de messagerie</span><span class="sxs-lookup"><span data-stu-id="7b451-106">**Add-in type:** Content, Task pane, Mail</span></span>

## <a name="syntax"></a><span data-ttu-id="7b451-107">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="7b451-107">Syntax</span></span>

```XML
<SourceLocation DefaultValue="string" />
```

## <a name="contained-in"></a><span data-ttu-id="7b451-108">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="7b451-108">Contained in</span></span>

- <span data-ttu-id="7b451-109">[DefaultSettings](defaultsettings.md) (compléments de contenu et de volet Office)</span><span class="sxs-lookup"><span data-stu-id="7b451-109">[DefaultSettings](defaultsettings.md) (Content and task pane add-ins)</span></span>
- <span data-ttu-id="7b451-110">[FormSettings](formsettings.md) (compléments de messagerie)</span><span class="sxs-lookup"><span data-stu-id="7b451-110">[FormSettings](formsettings.md) (Mail add-ins)</span></span>
- <span data-ttu-id="7b451-111">[ExtensionPoint](extensionpoint.md) (contextuel et LaunchEvent (aperçu) des compléments de messagerie)</span><span class="sxs-lookup"><span data-stu-id="7b451-111">[ExtensionPoint](extensionpoint.md) (Contextual and LaunchEvent (preview) mail add-ins)</span></span>

## <a name="can-contain"></a><span data-ttu-id="7b451-112">Peut contenir</span><span class="sxs-lookup"><span data-stu-id="7b451-112">Can contain</span></span>

[<span data-ttu-id="7b451-113">Override</span><span class="sxs-lookup"><span data-stu-id="7b451-113">Override</span></span>](override.md)

## <a name="attributes"></a><span data-ttu-id="7b451-114">Attributs</span><span class="sxs-lookup"><span data-stu-id="7b451-114">Attributes</span></span>

|<span data-ttu-id="7b451-115">**Attribut**</span><span class="sxs-lookup"><span data-stu-id="7b451-115">**Attribute**</span></span>|<span data-ttu-id="7b451-116">**Type**</span><span class="sxs-lookup"><span data-stu-id="7b451-116">**Type**</span></span>|<span data-ttu-id="7b451-117">**Obligatoire**</span><span class="sxs-lookup"><span data-stu-id="7b451-117">**Required**</span></span>|<span data-ttu-id="7b451-118">**Description**</span><span class="sxs-lookup"><span data-stu-id="7b451-118">**Description**</span></span>|
|:-----|:-----|:-----|:-----|
|<span data-ttu-id="7b451-119">DefaultValue</span><span class="sxs-lookup"><span data-stu-id="7b451-119">DefaultValue</span></span>|<span data-ttu-id="7b451-120">URL</span><span class="sxs-lookup"><span data-stu-id="7b451-120">URL</span></span>|<span data-ttu-id="7b451-121">obligatoire</span><span class="sxs-lookup"><span data-stu-id="7b451-121">required</span></span>|<span data-ttu-id="7b451-122">Spécifie la valeur par défaut de ce paramètre pour les paramètres régionaux spécifiés dans l’élément [DefaultLocale](defaultlocale.md).</span><span class="sxs-lookup"><span data-stu-id="7b451-122">Specifies the default value for this setting for the locale specified in the [DefaultLocale](defaultlocale.md) element.</span></span>|
