---
title: Élément AllowSnapshot dans le fichier manifeste
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: 02d44167dd1fd46ec6316f3e04393c99f19c9ff0
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450673"
---
# <a name="allowsnapshot-element"></a><span data-ttu-id="81bcc-102">AllowSnapshot, élément</span><span class="sxs-lookup"><span data-stu-id="81bcc-102">AllowSnapshot element</span></span>

<span data-ttu-id="81bcc-103">Indique si une capture instantanée de votre complément de contenu est enregistrée avec le document hôte.</span><span class="sxs-lookup"><span data-stu-id="81bcc-103">Specifies whether a snapshot image of your content add-in is saved with the host document.</span></span>

<span data-ttu-id="81bcc-104">**Type de complément :** Contenu</span><span class="sxs-lookup"><span data-stu-id="81bcc-104">**Add-in type:** Content</span></span>

## <a name="syntax"></a><span data-ttu-id="81bcc-105">Syntaxe</span><span class="sxs-lookup"><span data-stu-id="81bcc-105">Syntax</span></span>

```XML
<AllowSnapshot> [true | false]</AllowSnapshot>
```

## <a name="contained-in"></a><span data-ttu-id="81bcc-106">Contenu dans</span><span class="sxs-lookup"><span data-stu-id="81bcc-106">Contained in</span></span>

[<span data-ttu-id="81bcc-107">OfficeApp</span><span class="sxs-lookup"><span data-stu-id="81bcc-107">OfficeApp</span></span>](officeapp.md)

## <a name="remarks"></a><span data-ttu-id="81bcc-108">Remarques</span><span class="sxs-lookup"><span data-stu-id="81bcc-108">Remarks</span></span>

 > [!IMPORTANT]
 > <span data-ttu-id="81bcc-109">**AllowSnapshot** est défini sur `true` par défaut.</span><span class="sxs-lookup"><span data-stu-id="81bcc-109">**AllowSnapshot** is `true` by default.</span></span> <span data-ttu-id="81bcc-110">Cela crée une image du complément visible pour les utilisateurs qui ouvrent le document dans une version de l’application hôte ne prenant pas en charge les compléments Office, ou fournissant une image statique du complément si l’application hôte ne peut pas se connecter au serveur qui héberge le complément.</span><span class="sxs-lookup"><span data-stu-id="81bcc-110">This makes an image of the add-in visible for users that open the document in a version of the host application that doesn't support Office Add-ins, or provides a static image of the add-in if the host application can't connect to the server hosting the add-in.</span></span> <span data-ttu-id="81bcc-111">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span><span class="sxs-lookup"><span data-stu-id="81bcc-111">However, this also means that potentially sensitive information displayed in the add-in can be accessed directly from the document hosting the add-in.</span></span>

