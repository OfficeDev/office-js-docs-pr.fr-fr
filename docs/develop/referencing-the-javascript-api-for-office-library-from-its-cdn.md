---
title: Référencement de la bibliothèque de l’API JavaScript Office
description: Découvrez comment référencer la bibliothèque d’API JavaScript Office et les définitions de type dans votre complément.
ms.date: 02/27/2020
localization_priority: Normal
ms.openlocfilehash: 3b468ae9286eb53ffd5d3c0199c14045131105f3
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719349"
---
# <a name="referencing-the-office-javascript-api-library"></a><span data-ttu-id="99665-103">Référencement de la bibliothèque de l’API JavaScript Office</span><span class="sxs-lookup"><span data-stu-id="99665-103">Referencing the Office JavaScript API library</span></span>

<span data-ttu-id="99665-104">La bibliothèque de l' [API JavaScript pour Office](../reference/javascript-api-for-office.md) fournit les API que votre complément peut utiliser pour interagir avec l’hôte Office.</span><span class="sxs-lookup"><span data-stu-id="99665-104">The [Office JavaScript API](../reference/javascript-api-for-office.md) library provides the APIs that your add-in can use to interact with the Office host.</span></span> <span data-ttu-id="99665-105">Pour référencer la bibliothèque, le moyen le plus simple consiste à utiliser le réseau de distribution de contenu (CDN `<script>` ) en ajoutant `<head>` la balise suivante dans la section de votre page HTML :</span><span class="sxs-lookup"><span data-stu-id="99665-105">The simplest way to reference the library is to use the content delivery network (CDN) by adding the following `<script>` tag within the `<head>` section of your HTML page:</span></span>  

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
</head>
```

<span data-ttu-id="99665-106">Cela permet de télécharger et de mettre en cache les fichiers de l’API JavaScript pour Office la première fois que votre complément se charge pour s’assurer qu’il utilise l’implémentation la plus récente d’Office. js et ses fichiers associés pour la version spécifiée.</span><span class="sxs-lookup"><span data-stu-id="99665-106">This will download and cache the Office JavaScript API files the first time your add-in loads to make sure that it is using the most up-to-date implementation of Office.js and its associated files for the specified version.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="99665-107">Vous devez référencer l’API JavaScript Office depuis l' `<head>` intérieur de la section de la page pour vérifier que l’API est entièrement initialisée avant tout élément Body.</span><span class="sxs-lookup"><span data-stu-id="99665-107">You must reference the Office JavaScript API from inside the `<head>` section of the page to ensure that the API is fully initialized prior to any body elements.</span></span> <span data-ttu-id="99665-108">Les hôtes Office exigent que les compléments soient initialisés 5 secondes après l’activation.</span><span class="sxs-lookup"><span data-stu-id="99665-108">Office hosts require that add-ins initialize within 5 seconds of activation.</span></span> <span data-ttu-id="99665-109">Si votre complément n’est pas activé dans ce délai, il sera déclaré comme bloqué et un message d’erreur sera affiché à l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="99665-109">If your add-in doesn't activate within this threshold, it will be declared unresponsive and an error message will be displayed to the user.</span></span>

## <a name="api-versioning-and-backward-compatibility"></a><span data-ttu-id="99665-110">Contrôle de version de l’API et compatibilité descendante</span><span class="sxs-lookup"><span data-stu-id="99665-110">API versioning and backward compatibility</span></span>

<span data-ttu-id="99665-111">Dans l’extrait de code HTML précédent `/1/` , l’élément `office.js` devant dans l’URL du CDN spécifie la dernière version incrémentielle au sein de la version 1 d’Office. js.</span><span class="sxs-lookup"><span data-stu-id="99665-111">In the previous HTML snippet, the `/1/` in front of `office.js` in the CDN URL specifies the latest incremental release within version 1 of Office.js.</span></span> <span data-ttu-id="99665-112">Étant donné que l’API JavaScript pour Office conserve la compatibilité descendante, la dernière version continuera à prendre en charge les membres d’API qui ont été introduits précédemment dans la version 1.</span><span class="sxs-lookup"><span data-stu-id="99665-112">Because the Office JavaScript API maintains backward compatibility, the latest release will continue to support API members that were introduced earlier in version 1.</span></span> <span data-ttu-id="99665-113">Si vous devez mettre à niveau un projet existant, consultez [la rubrique mise à jour de la version de vos fichiers de schéma de manifeste et de l’API JavaScript pour Office](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span><span class="sxs-lookup"><span data-stu-id="99665-113">If you need to upgrade an existing project, see [Update the version of your Office JavaScript API and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span> 

<span data-ttu-id="99665-p104">Si vous envisagez de publier votre complément Office à partir d’AppSource, vous devez utiliser cette référence au CDN. Les références locales sont adaptées uniquement au développement interne et au débogage des scénarios.</span><span class="sxs-lookup"><span data-stu-id="99665-p104">If you plan to publish your Office Add-in from AppSource, you must use this CDN reference. Local references are only appropriate for internal, development, and debugging scenarios.</span></span>

> [!NOTE]
> <span data-ttu-id="99665-116">Pour utiliser les API destinées à la prévisualisation, référencez la version d’évaluation de la bibliothèque de l’interface API JavaScript Office dans le CDN : `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.</span><span class="sxs-lookup"><span data-stu-id="99665-116">To use preview APIs, reference the preview version of the Office JavaScript API library on the CDN: `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.</span></span>

## <a name="enabling-intellisense-for-a-typescript-project"></a><span data-ttu-id="99665-117">Activation d’IntelliSense pour un projet de dactylographié</span><span class="sxs-lookup"><span data-stu-id="99665-117">Enabling Intellisense for a TypeScript project</span></span>

<span data-ttu-id="99665-118">En plus de référencer l’API JavaScript pour Office, comme décrit précédemment, vous pouvez également activer IntelliSense pour le projet de complément de récriture à l’aide des définitions de type de [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js).</span><span class="sxs-lookup"><span data-stu-id="99665-118">In addition to referencing the Office JavaScript API as described previously, you can also enable Intellisense for TypeScript add-in project by using the type definitions from [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js).</span></span> <span data-ttu-id="99665-119">Pour ce faire, exécutez la commande suivante dans une invite du système à nœud (ou une fenêtre git bash) à partir de la racine de votre dossier de projet.</span><span class="sxs-lookup"><span data-stu-id="99665-119">To do so, run the following command in a Node-enabled system prompt (or git bash window) from the root of your project folder.</span></span> <span data-ttu-id="99665-120">[Node.js](https://nodejs.org) doit être installé (qui inclut npm).</span><span class="sxs-lookup"><span data-stu-id="99665-120">You must have [Node.js](https://nodejs.org) installed (which includes npm).</span></span>

```command&nbsp;line
npm install --save-dev @types/office-js
```

> [!NOTE]
> <span data-ttu-id="99665-121">Pour activer IntelliSense pour les API d’aperçu, utilisez les définitions de type d’aperçu de [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js-preview) en exécutant la commande suivante dans la racine du dossier de votre projet :</span><span class="sxs-lookup"><span data-stu-id="99665-121">To enable Intellisense for preview APIs, use the preview type definitions from [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js-preview) by running the following command in the root of your project folder:</span></span> 
>
> `npm install --save-dev @types/office-js-preview`

## <a name="see-also"></a><span data-ttu-id="99665-122">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="99665-122">See also</span></span>

- [<span data-ttu-id="99665-123">Compréhension de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="99665-123">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="99665-124">API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="99665-124">Office JavaScript API</span></span>](../reference/javascript-api-for-office.md)
