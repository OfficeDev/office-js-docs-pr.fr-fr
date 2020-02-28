---
title: Référencement de la bibliothèque d’API JavaScript pour Office
description: Découvrez comment référencer la bibliothèque d’API JavaScript Office et les définitions de type dans votre complément.
ms.date: 02/27/2020
localization_priority: Normal
ms.openlocfilehash: 9f7753b24e0a5861778b09ea93fecdc26fd2ca96
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325156"
---
# <a name="referencing-the-office-javascript-api-library"></a>Référencement de la bibliothèque d’API JavaScript pour Office

La bibliothèque de l' [API JavaScript pour Office](../reference/javascript-api-for-office.md) fournit les API que votre complément peut utiliser pour interagir avec l’hôte Office. Pour référencer la bibliothèque, le moyen le plus simple consiste à utiliser le réseau de distribution de contenu (CDN `<script>` ) en ajoutant `<head>` la balise suivante dans la section de votre page HTML :  

```html
<head>
    ...
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
</head>
```

Cela permet de télécharger et de mettre en cache les fichiers de l’API JavaScript pour Office la première fois que votre complément se charge pour s’assurer qu’il utilise l’implémentation la plus récente d’Office. js et ses fichiers associés pour la version spécifiée.

> [!IMPORTANT]
> Vous devez référencer l’API JavaScript Office depuis l' `<head>` intérieur de la section de la page pour vérifier que l’API est entièrement initialisée avant tout élément Body. Les hôtes Office exigent que les compléments soient initialisés 5 secondes après l’activation. Si votre complément n’est pas activé dans ce délai, il sera déclaré comme bloqué et un message d’erreur sera affiché à l’utilisateur.

## <a name="api-versioning-and-backward-compatibility"></a>Contrôle de version de l’API et compatibilité descendante

Dans l’extrait de code HTML précédent `/1/` , l’élément `office.js` devant dans l’URL du CDN spécifie la dernière version incrémentielle au sein de la version 1 d’Office. js. Étant donné que l’API JavaScript pour Office conserve la compatibilité descendante, la dernière version continuera à prendre en charge les membres d’API qui ont été introduits précédemment dans la version 1. Si vous devez mettre à niveau un projet existant, consultez [la rubrique mise à jour de la version de vos fichiers de schéma de manifeste et de l’API JavaScript pour Office](update-your-javascript-api-for-office-and-manifest-schema-version.md). 

Si vous envisagez de publier votre complément Office à partir d’AppSource, vous devez utiliser cette référence au CDN. Les références locales sont adaptées uniquement au développement interne et au débogage des scénarios.

> [!NOTE]
> Pour utiliser les API d’aperçu, référencez la version d’évaluation de la bibliothèque de l’API `https://appsforoffice.microsoft.com/lib/beta/hosted/office.js`JavaScript pour Office sur le CDN :.

## <a name="enabling-intellisense-for-a-typescript-project"></a>Activation d’IntelliSense pour un projet de dactylographié

En plus de référencer l’API JavaScript pour Office, comme décrit précédemment, vous pouvez également activer IntelliSense pour le projet de complément de récriture à l’aide des définitions de type de [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js). Pour ce faire, exécutez la commande suivante dans une invite du système à nœud (ou une fenêtre git bash) à partir de la racine de votre dossier de projet. [Node.js](https://nodejs.org) doit être installé (qui inclut npm).

```command&nbsp;line
npm install --save-dev @types/office-js
```

> [!NOTE]
> Pour activer IntelliSense pour les API d’aperçu, utilisez les définitions de type d’aperçu de [DefinitelyTyped](https://github.com/DefinitelyTyped/DefinitelyTyped/tree/master/types/office-js-preview) en exécutant la commande suivante dans la racine du dossier de votre projet : 
>
> `npm install --save-dev @types/office-js-preview`

## <a name="see-also"></a>Voir aussi

- [Présentation de l’API JavaScript pour Office](understanding-the-javascript-api-for-office.md)
- [API JavaScript pour Office](/office/dev/add-ins/reference/javascript-api-for-office)
