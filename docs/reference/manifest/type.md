---
title: Élément type dans le fichier manifeste
description: ''
ms.date: 05/03/2019
localization_priority: Normal
ms.openlocfilehash: 1c053d65c5e3c6ce597c9912ec608e0b36bc623b
ms.sourcegitcommit: b3996b1444e520b44cf752e76eef50908386ca26
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/21/2019
ms.locfileid: "33628227"
---
# <a name="type-element"></a><span data-ttu-id="5aaf0-102">Élément Type</span><span class="sxs-lookup"><span data-stu-id="5aaf0-102">Type element</span></span>

<span data-ttu-id="5aaf0-103">Indique si le complément équivalent est un complément COM ou un XLL.</span><span class="sxs-lookup"><span data-stu-id="5aaf0-103">Specifies if the equivalent add-in is a COM addin or an XLL.</span></span>

<span data-ttu-id="5aaf0-104">**Type de complément :** Volet Office, fonction personnalisée</span><span class="sxs-lookup"><span data-stu-id="5aaf0-104">**Add-in type:** Task pane, Custom function</span></span>

## <a name="syntax"></a><span data-ttu-id="5aaf0-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="5aaf0-105">Syntax</span></span>

```XML
    <Type> [COM | XLL] </Type>  
```

## <a name="contained-in"></a><span data-ttu-id="5aaf0-106">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="5aaf0-106">Contained in</span></span>

[<span data-ttu-id="5aaf0-107">EquivalentAdd-in</span><span class="sxs-lookup"><span data-stu-id="5aaf0-107">EquivalentAdd-in</span></span>](equivalentaddin.md)

## <a name="add-in-type-values"></a><span data-ttu-id="5aaf0-108">Valeurs de type de complément</span><span class="sxs-lookup"><span data-stu-id="5aaf0-108">Add-in type values</span></span>

<span data-ttu-id="5aaf0-109">Vous devez spécifier l’une des valeurs suivantes pour l' `Type` élément.</span><span class="sxs-lookup"><span data-stu-id="5aaf0-109">You must specify one of the following values for the `Type` element.</span></span>

- <span data-ttu-id="5aaf0-110">COM : spécifie que le complément équivalent est un complément COM.</span><span class="sxs-lookup"><span data-stu-id="5aaf0-110">COM: Specifies the equivalent add-in is a COM add-in.</span></span>
- <span data-ttu-id="5aaf0-111">XLL : spécifie que le complément équivalent est une XLL Excel.</span><span class="sxs-lookup"><span data-stu-id="5aaf0-111">XLL: Specifies the equivalent add-in is an Excel XLL.</span></span>

## <a name="see-also"></a><span data-ttu-id="5aaf0-112">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="5aaf0-112">See also</span></span>

- [<span data-ttu-id="5aaf0-113">Rendre vos fonctions personnalisées compatibles avec les fonctions XLL définies par l’utilisateur</span><span class="sxs-lookup"><span data-stu-id="5aaf0-113">Make your custom functions compatible with XLL user-defined functions</span></span>](../../excel/make-custom-functions-compatible-with-xll-udf.md)
- [<span data-ttu-id="5aaf0-114">Faire en sorte que votre complément Excel soit compatible avec un complément COM existant</span><span class="sxs-lookup"><span data-stu-id="5aaf0-114">Make your Excel add-in compatible with an existing COM add-in</span></span>](../../develop/make-office-add-in-compatible-with-existing-com-add-in.md)