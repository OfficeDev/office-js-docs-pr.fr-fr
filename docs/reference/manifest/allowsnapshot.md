---
title: Élément AllowSnapshot dans le fichier manifeste
description: Indique si une capture instantanée de votre complément de contenu est enregistrée avec le document hôte.
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 8bb143d13a17b3e184af64f1bf18f2a32a55b60c
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720959"
---
# <a name="allowsnapshot-element"></a><span data-ttu-id="30d89-103">AllowSnapshot, élément</span><span class="sxs-lookup"><span data-stu-id="30d89-103">AllowSnapshot element</span></span>

<span data-ttu-id="30d89-104">Indique si une capture instantanée de votre complément de contenu est enregistrée avec le document hôte.</span><span class="sxs-lookup"><span data-stu-id="30d89-104">Specifies whether a snapshot image of your content add-in is saved with the host document.</span></span>

<span data-ttu-id="30d89-105">**Type de complément :** Contenu</span><span class="sxs-lookup"><span data-stu-id="30d89-105">**Add-in type:** Content</span></span>

## <a name="syntax"></a><span data-ttu-id="30d89-106">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="30d89-106">Syntax</span></span>

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a><span data-ttu-id="30d89-107">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="30d89-107">Contained in</span></span>

[<span data-ttu-id="30d89-108">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="30d89-108">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="30d89-109">Remarques</span><span class="sxs-lookup"><span data-stu-id="30d89-109">Remarks</span></span>

 > [!IMPORTANT]
 > <span data-ttu-id="30d89-110">**AllowSnapshot** est défini sur `true` par défaut.</span><span class="sxs-lookup"><span data-stu-id="30d89-110">**AllowSnapshot** is `true` by default.</span></span> <span data-ttu-id="30d89-111">Cela crée une image du complément visible pour les utilisateurs qui ouvrent le document dans une version de l’application hôte ne prenant pas en charge les compléments Office, ou fournissant une image statique du complément si l’application hôte ne peut pas se connecter au serveur qui héberge le complément.</span><span class="sxs-lookup"><span data-stu-id="30d89-111">This makes an image of the add-in visible for users that open the document in a version of the host application that doesn't support Office Add-ins, or provides a static image of the add-in if the host application can't connect to the server hosting the add-in.</span></span> <span data-ttu-id="30d89-112">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span><span class="sxs-lookup"><span data-stu-id="30d89-112">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span></span>

