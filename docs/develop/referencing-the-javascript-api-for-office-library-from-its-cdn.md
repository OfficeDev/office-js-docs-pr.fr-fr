---
title: Référencement de la bibliothèque de l’API JavaScript Office
description: Découvrez comment référencer la bibliothèque d’API JavaScript Office et les définitions de type dans votre complément.
ms.date: 06/23/2020
localization_priority: Normal
ms.openlocfilehash: 3f90b0798b14b66fe6d01f62eca3802fce179bec
ms.sourcegitcommit: a4873c3525c7d30ef551545d27eb2c0a16b4eb50
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/25/2020
ms.locfileid: "44888130"
---
# <a name="referencing-the-office-javascript-api-library"></a>Référencement de la bibliothèque de l’API JavaScript Office

La bibliothèque de l' [API JavaScript pour Office](../reference/javascript-api-for-office.md) fournit les API que votre complément peut utiliser pour interagir avec l’hôte Office. Pour référencer la bibliothèque, le moyen le plus simple consiste à utiliser le réseau de distribution de contenu (CDN) en ajoutant la `<script>` balise suivante dans la `<head>` section de votre page HTML :  

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
</head>
```

Cela permet de télécharger et de mettre en cache les fichiers de l’API JavaScript pour Office la première fois que votre complément se charge pour s’assurer qu’il utilise l’implémentation la plus à jour de Office.js et ses fichiers associés pour la version spécifiée.

> [!IMPORTANT]
> Vous devez référencer l’API JavaScript Office depuis l’intérieur `<head>` de la section de la page pour vérifier que l’API est entièrement initialisée avant tout élément Body. Les hôtes Office exigent que les compléments soient initialisés 5 secondes après l’activation. Si votre complément n’est pas activé dans ce délai, il sera déclaré comme bloqué et un message d’erreur sera affiché à l’utilisateur.

## <a name="api-versioning-and-backward-compatibility"></a>Contrôle de version de l’API et compatibilité descendante

Dans l’extrait de code HTML précédent, l’élément `/1/` devant `office.js` dans l’URL du CDN spécifie la dernière version incrémentielle dans la version 1 de Office.js. Étant donné que l’API JavaScript pour Office conserve la compatibilité descendante, la dernière version continuera à prendre en charge les membres d’API qui ont été introduits précédemment dans la version 1. Si vous devez mettre à niveau un projet existant, consultez [la rubrique mise à jour de la version de vos fichiers de schéma de manifeste et de l’API JavaScript pour Office](update-your-javascript-api-for-office-and-manifest-schema-version.md). 

If you plan to publish your Office Add-in from AppSource, you must use this CDN reference. Local references are only appropriate for internal, development, and debugging scenarios.

> [!NOTE]
> Pour utiliser les API destinées à la prévisualisation, référencez la version d’évaluation de la bibliothèque de l’interface API JavaScript Office dans le CDN : `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`.

## <a name="enabling-intellisense-for-a-typescript-project"></a>Activation d’IntelliSense pour un projet de dactylographié

En plus de référencer l’API JavaScript pour Office, comme décrit précédemment, vous pouvez également activer IntelliSense pour le projet de complément de récriture à l’aide des définitions de type de [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js). Pour ce faire, exécutez la commande suivante dans une invite du système à nœud (ou une fenêtre git bash) à partir de la racine de votre dossier de projet. [Node.js](https://nodejs.org) doit être installé (qui inclut npm).

```command&nbsp;line
npm install --save-dev @types/office-js
```

## <a name="preview-apis"></a>API d’aperçu

De nouvelles API JavaScript sont introduites pour la première fois dans « Preview », puis elles deviennent une partie d’un ensemble de conditions requises spécifiques, après un test suffisant, et les commentaires des utilisateurs sont nécessaires.

[!INCLUDE [Information about using preview APIs](../includes/using-preview-apis-host.md)]

## <a name="see-also"></a>Voir aussi

- [Compréhension de l’API JavaScript pour Office](understanding-the-javascript-api-for-office.md)
- [API JavaScript pour Office](../reference/javascript-api-for-office.md)
