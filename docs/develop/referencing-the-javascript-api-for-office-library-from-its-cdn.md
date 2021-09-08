---
title: Référencement de la bibliothèque de l’API JavaScript Office
description: Découvrez comment référencer la bibliothèque Office’API JavaScript et les définitions de type dans votre application.
ms.date: 02/18/2021
localization_priority: Normal
ms.openlocfilehash: 04f97412c07cb39f5b2f753c3ce14e56e87c3de5
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936457"
---
# <a name="referencing-the-office-javascript-api-library"></a>Référencement de la bibliothèque de l’API JavaScript Office

La [Office’API JavaScript](../reference/javascript-api-for-office.md) fournit les API que votre application peut utiliser pour interagir avec l’application Office web. Le moyen le plus simple de référencer la bibliothèque consiste à utiliser le réseau de distribution de contenu (CDN) en ajoutant la balise suivante dans la `<script>` section de votre page `<head>` HTML.

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
</head>
```

Cela permet de télécharger et de mettre en cache les fichiers de l’API JavaScript Office la première fois que votre application se charge pour s’assurer qu’elle utilise l’implémentation la plus à jour de Office.js et de ses fichiers associés pour la version spécifiée.

> [!IMPORTANT]
> Vous devez référencer l’API JavaScript Office à partir de la section de la page pour vous assurer que l’API est entièrement initialisée avant les éléments `<head>` body.

## <a name="api-versioning-and-backward-compatibility"></a>Gestion des versions d’API et compatibilité avec les versions antérieures

Dans l’extrait de code HTML précédent, l’URL avant de l’CDN spécifie la dernière version incrémentielle dans la version 1 de `/1/` `office.js` Office.js. Étant donné que Office API JavaScript maintient la compatibilité ascendante, la dernière version continuera à prendre en charge les membres d’API qui ont été introduits précédemment dans la version 1. Si vous avez besoin de mettre à niveau un projet existant, voir Mettre à jour la version de votre API JavaScript Office fichiers de schéma de manifeste et de [l’API JavaScript.](update-your-javascript-api-for-office-and-manifest-schema-version.md) 

Si vous envisagez de publier votre complément Office à partir d’AppSource, vous devez utiliser cette référence au CDN. Les références locales sont adaptées uniquement au développement interne et au débogage des scénarios.

> [!NOTE]
> Pour utiliser les API destinées à la prévisualisation, référencez la version d’évaluation de la bibliothèque de l’interface API JavaScript Office dans le CDN : `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.

## <a name="enabling-intellisense-for-a-typescript-project"></a>Activation IntelliSense pour un projet TypeScript

En plus de référencer l’API JavaScript Office comme décrit précédemment, vous pouvez également activer IntelliSense pour le projet de complément TypeScript à l’aide des définitions de type de [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js). Pour ce faire, exécutez la commande suivante dans une invite système node (ou une fenêtre Git Bash) à partir de la racine du dossier de votre projet. [Node.js](https://nodejs.org) doit être installé (qui inclut npm).

```command&nbsp;line
npm install --save-dev @types/office-js
```

## <a name="preview-apis"></a>API d’aperçu

Les nouvelles API JavaScript sont d’abord introduites dans « aperçu », puis font partie d’un ensemble de conditions requises numérotées spécifique une fois que des tests suffisants ont eu lieu et que des commentaires de l’utilisateur sont requis.

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

## <a name="see-also"></a>Voir aussi

- [Compréhension de l’API JavaScript pour Office](understanding-the-javascript-api-for-office.md)
- [API JavaScript pour Office](../reference/javascript-api-for-office.md)
