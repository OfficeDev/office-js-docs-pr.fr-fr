---
title: Référencement de la bibliothèque de l’API JavaScript Office
description: Découvrez comment référencer la bibliothèque d’API JavaScript Office et les définitions de type dans votre complément.
ms.date: 02/18/2021
ms.localizationpriority: medium
ms.openlocfilehash: 38121fe3d3df0a86fef3e2c8e3a58399640f1e2a
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/06/2022
ms.locfileid: "66660115"
---
# <a name="referencing-the-office-javascript-api-library"></a>Référencement de la bibliothèque de l’API JavaScript Office

La bibliothèque [d’API JavaScript Office](../reference/javascript-api-for-office.md) fournit les API que votre complément peut utiliser pour interagir avec l’application Office. Le moyen le plus simple de référencer la bibliothèque consiste à utiliser le réseau de distribution de contenu (CDN) en ajoutant la balise suivante `<script>` dans la `<head>` section de votre page HTML.

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
</head>
```

Cela permet de télécharger et de mettre en cache les fichiers d’API JavaScript Office la première fois que votre complément se charge pour s’assurer qu’il utilise l’implémentation la plus récente de Office.js et ses fichiers associés pour la version spécifiée.

> [!IMPORTANT]
> Vous devez référencer l’API JavaScript Office à partir de la `<head>` section de la page pour vous assurer que l’API est entièrement initialisée avant tous les éléments de corps.

## <a name="api-versioning-and-backward-compatibility"></a>Contrôle de version d’API et compatibilité descendante

Dans l’extrait de code HTML précédent, le `/1/` devant de `office.js` l’URL CDN spécifie la dernière version incrémentielle dans la version 1 de Office.js. Étant donné que l’API JavaScript Office conserve une compatibilité descendante, la dernière version continuera de prendre en charge les membres de l’API qui ont été introduits précédemment dans la version 1. Si vous devez mettre à niveau un projet existant, consultez [Mettre à jour la version de votre API JavaScript Office et les fichiers de schéma de manifeste](update-your-javascript-api-for-office-and-manifest-schema-version.md). 

Si vous envisagez de publier votre complément Office à partir d’AppSource, vous devez utiliser cette référence au CDN. Les références locales sont adaptées uniquement au développement interne et au débogage des scénarios.

> [!NOTE]
> Pour utiliser les API destinées à la prévisualisation, référencez la version d’évaluation de la bibliothèque de l’interface API JavaScript Office dans le CDN : `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.

## <a name="enabling-intellisense-for-a-typescript-project"></a>Activation d’IntelliSense pour un projet TypeScript

En plus de référencer l’API JavaScript Office comme décrit précédemment, vous pouvez également activer IntelliSense pour le projet de complément TypeScript à l’aide des définitions de type de [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js). Pour ce faire, exécutez la commande suivante dans une invite système prenant en charge node (ou une fenêtre git bash) à partir de la racine de votre dossier de projet. [Node.js](https://nodejs.org) doit être installé (qui inclut npm).

```command&nbsp;line
npm install --save-dev @types/office-js
```

## <a name="preview-apis"></a>API d’aperçu

Les nouvelles API JavaScript sont introduites en « préversion » et font ensuite partie d’un ensemble spécifique de conditions requises numérotées après que des tests suffisants se produisent et que les commentaires des utilisateurs sont acquis.

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

## <a name="see-also"></a>Voir aussi

- [Compréhension de l’API JavaScript pour Office](understanding-the-javascript-api-for-office.md)
- [API JavaScript pour Office](../reference/javascript-api-for-office.md)
