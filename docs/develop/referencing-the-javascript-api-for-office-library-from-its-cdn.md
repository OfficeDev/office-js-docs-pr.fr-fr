---
title: Référencement de la bibliothèque de l’API JavaScript Office
description: Découvrez comment référencer la bibliothèque de l’API JavaScript office et les définitions de type dans votre application.
ms.date: 02/18/2021
localization_priority: Normal
ms.openlocfilehash: 346a34c0cbc31b5e569a5106dcd2bc01593b114a
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505191"
---
# <a name="referencing-the-office-javascript-api-library"></a><span data-ttu-id="3dcba-103">Référencement de la bibliothèque de l’API JavaScript Office</span><span class="sxs-lookup"><span data-stu-id="3dcba-103">Referencing the Office JavaScript API library</span></span>

<span data-ttu-id="3dcba-104">La [bibliothèque de l’API JavaScript](../reference/javascript-api-for-office.md) pour Office fournit les API que votre application peut utiliser pour interagir avec l’application Office.</span><span class="sxs-lookup"><span data-stu-id="3dcba-104">The [Office JavaScript API](../reference/javascript-api-for-office.md) library provides the APIs that your add-in can use to interact with the Office application.</span></span> <span data-ttu-id="3dcba-105">Le moyen le plus simple de référencer la bibliothèque consiste à utiliser le réseau de distribution de contenu (CDN) en ajoutant la balise suivante dans la `<script>` section de votre page HTML `<head>` :</span><span class="sxs-lookup"><span data-stu-id="3dcba-105">The simplest way to reference the library is to use the content delivery network (CDN) by adding the following `<script>` tag within the `<head>` section of your HTML page:</span></span>  

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
</head>
```

<span data-ttu-id="3dcba-106">Cela permet de télécharger et de mettre en cache les fichiers de l’API JavaScript pour Office lors du premier chargement de votre application pour vous assurer qu’elle utilise l’implémentation la plus à jour de Office.js et de ses fichiers associés pour la version spécifiée.</span><span class="sxs-lookup"><span data-stu-id="3dcba-106">This will download and cache the Office JavaScript API files the first time your add-in loads to make sure that it is using the most up-to-date implementation of Office.js and its associated files for the specified version.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3dcba-107">Vous devez référencer l’API JavaScript Office à partir de la section de la page pour vous assurer que l’API est entièrement initialisée avant les `<head>` éléments body.</span><span class="sxs-lookup"><span data-stu-id="3dcba-107">You must reference the Office JavaScript API from inside the `<head>` section of the page to ensure that the API is fully initialized prior to any body elements.</span></span>

## <a name="api-versioning-and-backward-compatibility"></a><span data-ttu-id="3dcba-108">Gestion des versions d’API et compatibilité avec les versions antérieures</span><span class="sxs-lookup"><span data-stu-id="3dcba-108">API versioning and backward compatibility</span></span>

<span data-ttu-id="3dcba-109">Dans l’extrait de code HTML précédent, la partie avant de l’URL du CDN spécifie la dernière version incrémentielle dans la version 1 de `/1/` `office.js` Office.js.</span><span class="sxs-lookup"><span data-stu-id="3dcba-109">In the previous HTML snippet, the `/1/` in front of `office.js` in the CDN URL specifies the latest incremental release within version 1 of Office.js.</span></span> <span data-ttu-id="3dcba-110">Étant donné que l’API JavaScript Pour Office maintient la compatibilité ascendante, la dernière version continuera à prendre en charge les membres d’API qui ont été introduits précédemment dans la version 1.</span><span class="sxs-lookup"><span data-stu-id="3dcba-110">Because the Office JavaScript API maintains backward compatibility, the latest release will continue to support API members that were introduced earlier in version 1.</span></span> <span data-ttu-id="3dcba-111">Si vous devez mettre à niveau un projet existant, voir Mettre à jour la version de votre API JavaScript Office et des fichiers [de schéma de manifeste.](update-your-javascript-api-for-office-and-manifest-schema-version.md)</span><span class="sxs-lookup"><span data-stu-id="3dcba-111">If you need to upgrade an existing project, see [Update the version of your Office JavaScript API and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span> 

<span data-ttu-id="3dcba-p103">Si vous envisagez de publier votre complément Office à partir d’AppSource, vous devez utiliser cette référence au CDN. Les références locales sont adaptées uniquement au développement interne et au débogage des scénarios.</span><span class="sxs-lookup"><span data-stu-id="3dcba-p103">If you plan to publish your Office Add-in from AppSource, you must use this CDN reference. Local references are only appropriate for internal, development, and debugging scenarios.</span></span>

> [!NOTE]
> <span data-ttu-id="3dcba-114">Pour utiliser les API destinées à la prévisualisation, référencez la version d’évaluation de la bibliothèque de l’interface API JavaScript Office dans le CDN : `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.</span><span class="sxs-lookup"><span data-stu-id="3dcba-114">To use preview APIs, reference the preview version of the Office JavaScript API library on the CDN: `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.</span></span>

## <a name="enabling-intellisense-for-a-typescript-project"></a><span data-ttu-id="3dcba-115">Activation IntelliSense pour un projet TypeScript</span><span class="sxs-lookup"><span data-stu-id="3dcba-115">Enabling IntelliSense for a TypeScript project</span></span>

<span data-ttu-id="3dcba-116">En plus de référencer l’API JavaScript Office comme décrit précédemment, vous pouvez également activer IntelliSense pour le projet de complément TypeScript à l’aide des définitions de type de [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js).</span><span class="sxs-lookup"><span data-stu-id="3dcba-116">In addition to referencing the Office JavaScript API as described previously, you can also enable IntelliSense for TypeScript add-in project by using the type definitions from [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js).</span></span> <span data-ttu-id="3dcba-117">Pour ce faire, exécutez la commande suivante dans une invite système node (ou une fenêtre Git Bash) à partir de la racine du dossier de votre projet.</span><span class="sxs-lookup"><span data-stu-id="3dcba-117">To do so, run the following command in a Node-enabled system prompt (or git bash window) from the root of your project folder.</span></span> <span data-ttu-id="3dcba-118">[Node.js](https://nodejs.org) doit être installé (qui inclut npm).</span><span class="sxs-lookup"><span data-stu-id="3dcba-118">You must have [Node.js](https://nodejs.org) installed (which includes npm).</span></span>

```command&nbsp;line
npm install --save-dev @types/office-js
```

## <a name="preview-apis"></a><span data-ttu-id="3dcba-119">API d’aperçu</span><span class="sxs-lookup"><span data-stu-id="3dcba-119">Preview APIs</span></span>

<span data-ttu-id="3dcba-120">Les nouvelles API JavaScript sont d’abord introduites dans « aperçu », puis font partie d’un ensemble de conditions requises numérotées spécifique une fois que des tests suffisants ont eu lieu et que des commentaires de l’utilisateur sont requis.</span><span class="sxs-lookup"><span data-stu-id="3dcba-120">New JavaScript APIs are first introduced in "preview" and later become part of a specific numbered requirement set after sufficient testing occurs and user feedback is required.</span></span>

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

## <a name="see-also"></a><span data-ttu-id="3dcba-121">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="3dcba-121">See also</span></span>

- [<span data-ttu-id="3dcba-122">Compréhension de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="3dcba-122">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="3dcba-123">API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="3dcba-123">Office JavaScript API</span></span>](../reference/javascript-api-for-office.md)
