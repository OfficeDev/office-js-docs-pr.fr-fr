---
title: Modèle d’objet JavaScript Word dans les compléments Office
description: Découvrez les classes les plus importantes dans le modèle objet JavaScript spécifique à Word.
ms.date: 10/14/2020
localization_priority: Priority
ms.openlocfilehash: c85c56987ef5de7c087064ac668f137326089642
ms.sourcegitcommit: 42e6cfe51d99d4f3f05a3245829d764b28c46bbb
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/23/2020
ms.locfileid: "48740867"
---
# <a name="word-javascript-object-model-in-office-add-ins"></a><span data-ttu-id="949ec-103">Modèle d’objet JavaScript Word dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="949ec-103">Word JavaScript object model in Office Add-ins</span></span>

<span data-ttu-id="949ec-104">Cet article décrit les concepts de base de l’utilisation de [l’API JavaScript pour Word](../reference/overview/word-add-ins-reference-overview.md) pour créer des compléments. Il présente les concepts fondamentaux de l’utilisation de l’API.</span><span class="sxs-lookup"><span data-stu-id="949ec-104">This article describes concepts that are fundamental to using the [Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) to build add-ins. It introduces core concepts that are fundamental to using the API.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="949ec-105">Pour en savoir plus sur la nature asynchrone des API Word et la manière dont elles fonctionnent avec le document, consultez [Utilisation du modèle d’API spécifique à l’application](../develop/application-specific-api-model.md).</span><span class="sxs-lookup"><span data-stu-id="949ec-105">See [Using the application-specific API model](../develop/application-specific-api-model.md) to learn about the asynchronous nature of the Word APIs and how they work with the document.</span></span>

## <a name="officejs-apis-for-word"></a><span data-ttu-id="949ec-106">API Office.js pour Word</span><span class="sxs-lookup"><span data-stu-id="949ec-106">Office.js APIs for Word</span></span>

<span data-ttu-id="949ec-107">Un complément Word interagit avec des objets dans Excel en utilisant l’API Office JavaScript, qui inclut deux modèles d’objets JavaScript :</span><span class="sxs-lookup"><span data-stu-id="949ec-107">A Word add-in interacts with objects in Excel by using the Office JavaScript API, which includes two JavaScript object models:</span></span>

* <span data-ttu-id="949ec-108">**API JavaScript Word** : l’[API JavaScript Word](../reference/overview/word-add-ins-reference-overview.md) fournit des objets fortement typés que vous pouvez utiliser pour accéder au document, à des plages, à des tableaux, à des listes, à une mise en forme, etc.</span><span class="sxs-lookup"><span data-stu-id="949ec-108">**Word JavaScript API**: The [Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) provides strongly-typed objects that you can use to access the document, ranges, tables, lists, formatting, and more.</span></span>

* <span data-ttu-id="949ec-109">**API communes** : l’[API commune](/javascript/api/office) peut être utilisée pour accéder à des fonctionnalités telles que l’interface utilisateur, les boîtes de dialogue et les paramètres de client communs à différents types d’applications Office.</span><span class="sxs-lookup"><span data-stu-id="949ec-109">**Common APIs**: The [Common API](/javascript/api/office) can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</span></span>

<span data-ttu-id="949ec-110">Vous utiliserez probablement l’API JavaScript Word pour développer la majorité des fonctionnalités des compléments destinés à Word, vous utiliserez également des objets dans l’API commune.</span><span class="sxs-lookup"><span data-stu-id="949ec-110">While you'll likely use the Word JavaScript API to develop the majority of functionality in add-ins that target Word, you'll also use objects in the Common API.</span></span> <span data-ttu-id="949ec-111">Par exemple :</span><span class="sxs-lookup"><span data-stu-id="949ec-111">For example:</span></span>

* <span data-ttu-id="949ec-112">[Context](/javascript/api/office/office.context) :le `Context` représente l’environnement d’exécution du complément et permet d’accéder à des objets clés de l’API.</span><span class="sxs-lookup"><span data-stu-id="949ec-112">[Context](/javascript/api/office/office.context): The `Context` object represents the runtime environment of the add-in and provides access to key objects of the API.</span></span> <span data-ttu-id="949ec-113">Il se compose de détails sur la configuration du document comme `contentLanguage` et `officeTheme`, et fournit des informations sur l’environnement d’exécution du complément comme `host` et `platform`.</span><span class="sxs-lookup"><span data-stu-id="949ec-113">It consists of document configuration details such as `contentLanguage` and `officeTheme` and also provides information about the add-in's runtime environment such as `host` and `platform`.</span></span> <span data-ttu-id="949ec-114">En outre, il fournit la méthode `requirements.isSetSupported()` que vous pouvez utiliser pour vérifier si un ensemble de conditions requises spécifié est pris en charge par l’application Excel dans laquelle le complément est exécuté.</span><span class="sxs-lookup"><span data-stu-id="949ec-114">Additionally, it provides the `requirements.isSetSupported()` method, which you can use to check whether a specified requirement set is supported by the Excel application where the add-in is running.</span></span>
* <span data-ttu-id="949ec-115">[Document](/javascript/api/office/office.document) : le `Document` fournit la méthode `getFileAsync()` que vous pouvez utiliser pour télécharger le fichier Word dans lequel le complément est exécuté.</span><span class="sxs-lookup"><span data-stu-id="949ec-115">[Document](/javascript/api/office/office.document): The `Document` object provides the `getFileAsync()` method, which you can use to download the Word file where the add-in is running.</span></span>

![Image des différences entre l’API Word et les API communes](../images/word-js-api-common-api.png)

## <a name="word-specific-object-model"></a><span data-ttu-id="949ec-117">Modèle d’objet spécifique à Word</span><span class="sxs-lookup"><span data-stu-id="949ec-117">Word-specific object model</span></span>

<span data-ttu-id="949ec-118">Pour comprendre les API Word, vous devez connaître la manière dont les composants d’un document sont liés les uns aux autres.</span><span class="sxs-lookup"><span data-stu-id="949ec-118">To understand the Word APIs, you must understand how the components of a document are related to one another.</span></span>

* <span data-ttu-id="949ec-119">Le **Document** contient les **Sections**, ainsi que les entités de niveau document telles que les paramètres et les parties XML personnalisées.</span><span class="sxs-lookup"><span data-stu-id="949ec-119">The **Document** contains the **Section**s, and document-level entities such as settings and custom XML parts.</span></span>
* <span data-ttu-id="949ec-120">Une **Section** contient un **Corps**.</span><span class="sxs-lookup"><span data-stu-id="949ec-120">A **Section** contains a **Body**.</span></span>
* <span data-ttu-id="949ec-121">Un **Corps** donne accès aux **Paragraphe**s, **ContentControl**s et **Plage** objets, entre autres.</span><span class="sxs-lookup"><span data-stu-id="949ec-121">A **Body** gives access to **Paragraph**s, **ContentControl**s, and **Range** objects, among others.</span></span>
* <span data-ttu-id="949ec-122">Une **Plage** représente une zone contiguë de contenu, y compris du texte, un espace vide, des **Tableaux** et des images.</span><span class="sxs-lookup"><span data-stu-id="949ec-122">A **Range** represents a contiguous area of content, including text, white space, **Table**s, and images.</span></span> <span data-ttu-id="949ec-123">Elle contient également la plupart des méthodes de manipulation de texte.</span><span class="sxs-lookup"><span data-stu-id="949ec-123">It also contains most of the text manipulation methods.</span></span>
* <span data-ttu-id="949ec-124">Une **Liste** représente le texte d’une liste numérotée ou une liste à puces.</span><span class="sxs-lookup"><span data-stu-id="949ec-124">A **List** represents text in a numbered or bulleted list.</span></span>

## <a name="see-also"></a><span data-ttu-id="949ec-125">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="949ec-125">See also</span></span>

- [<span data-ttu-id="949ec-126">Présentation de l’API JavaScript pour Word</span><span class="sxs-lookup"><span data-stu-id="949ec-126">Word JavaScript API overview</span></span>](../reference/overview/word-add-ins-reference-overview.md)
- [<span data-ttu-id="949ec-127">Créer votre premier complément Word</span><span class="sxs-lookup"><span data-stu-id="949ec-127">Build your first Word add-in</span></span>](../quickstarts/word-quickstart.md)
- [<span data-ttu-id="949ec-128">Didacticiel sur les compléments Word</span><span class="sxs-lookup"><span data-stu-id="949ec-128">Word add-in tutorial</span></span>](../tutorials/word-tutorial.md)
- [<span data-ttu-id="949ec-129">Référence sur l’API JavaScript pour Word</span><span class="sxs-lookup"><span data-stu-id="949ec-129">Word JavaScript API reference</span></span>](/javascript/api/word)
- [<span data-ttu-id="949ec-130">Découvrez le programme pour les développeurs Microsoft 365</span><span class="sxs-lookup"><span data-stu-id="949ec-130">Learn about the Microsoft 365 Developer Program</span></span>](https://developer.microsoft.com/microsoft-365/dev-program)