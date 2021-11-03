---
title: Configuration de votre environnement de développement
description: Configurer votre environnement de développement pour créer des Office de développement.
ms.date: 10/26/2021
ms.localizationpriority: medium
ms.openlocfilehash: 9dbe2a994dd8da028ecd1ae4a31b2c7847a062b1
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681173"
---
# <a name="set-up-your-development-environment"></a>Configuration de votre environnement de développement

Ce guide vous aide à configurer des outils pour créer des Office en suivant nos démarrages rapides ou didacticiels. Vous devez installer les outils à partir de la liste ci-dessous. Si vous avez déjà installé ces éléments, vous êtes prêt à commencer un démarrage rapide, tel que [celui-ci Excel React démarrage rapide.](../quickstarts/excel-quickstart-react.md)

- Node.js
- npm
- Un Microsoft 365 qui inclut la version d’abonnement de Office
- Éditeur de code de votre choix
- Le Office javascript de linter

Ce guide suppose que vous savez utiliser un outil de ligne de commande.

## <a name="install-nodejs"></a>Installer Node.js.

Node.js est un runtime JavaScript dont vous aurez besoin pour développer des Office modernes.

Installez Node.js en [téléchargeant la dernière version recommandée à partir de leur site web.](https://nodejs.org) Suivez les instructions d’installation de votre système d’exploitation.

## <a name="install-npm"></a>Installer npm

npm est un registre logiciel open source à partir duquel télécharger les packages utilisés dans le développement de Office de développement.

Pour installer npm, exécutez la commande suivante dans la ligne de commande.

```command&nbsp;line
    npm install npm -g
```

Pour vérifier si npm est déjà installé et voir la version installée, exécutez la commande suivante dans la ligne de commande.

```command&nbsp;line
npm -v
```

Vous pouvez utiliser un gestionnaire de version Node pour vous permettre de basculer entre plusieurs versions de Node.js et npm, mais cela n’est pas strictement nécessaire. Pour plus d’informations sur la façon de faire, voir [les instructions de npm.](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm)

## <a name="get-microsoft-365"></a>Obtenir Microsoft 365

Si vous n’avez pas encore de compte Microsoft 365, vous pouvez obtenir un abonnement Microsoft 365 renouvelable gratuit de 90 jours qui inclut toutes les applications Office en rejoignant le programme Microsoft 365 [développeur.](https://developer.microsoft.com/office/dev-program)

## <a name="install-a-code-editor"></a>Installer un éditeur de code

Vous pouvez utiliser n’importe quel éditeur de code ou IDE qui prend en charge le développement côté client pour créer votre composant WebPart, par exemple :

- [Visual Studio Code](https://code.visualstudio.com/)
- [Atom](https://atom.io)
- [Webstorm](https://www.jetbrains.com/webstorm)

## <a name="install-and-use-the-office-javascript-linter"></a>Installer et utiliser le Office JavaScript

Microsoft fournit un linter JavaScript pour vous aider à capturer les erreurs courantes lors de l’utilisation Office bibliothèque JavaScript. Pour installer le linter, exécutez les deux commandes suivantes (une fois que vous avez installé [Node.js](#install-nodejs) et [npm](#install-npm)).

```command&nbsp;line
npm install office-addin-lint --save-dev
npm install eslint-plugin-office-addins --save-dev
```

Si vous créez un projet de Office avec l’outil Yo Office, le reste de l’installation est terminé pour vous. Exécutez le linter avec la commande suivante dans le terminal d’un éditeur, par exemple Visual Studio Code, ou dans une invite de commandes. Les problèmes trouvés par le linter apparaissent dans le terminal ou l’invite, et apparaissent également directement dans le code lorsque vous utilisez un éditeur qui prend en charge les messages linter, tels que Visual Studio Code. (Pour plus d’informations sur l’installation de l’outil Yo Office, voir l’un de nos démarrages rapides de Office, comme [celui-ci](../quickstarts/excel-quickstart-jquery.md)pour les Excel.)

```command&nbsp;line
npm run lint
```

Si votre projet de add-in a été créé d’une autre façon, prenez les mesures suivantes.

1. À la racine du projet, créez un fichier texte nommé **.eslintrc.json,** s’il n’en existe pas déjà un. Assurez-vous qu’il possède des propriétés `plugins` nommées `extends` et , les deux types de tableau. Le `plugins` tableau doit inclure et le tableau doit inclure `"office-addins"` `extends` `"plugin:office-addins/recommended"` . Voici un exemple simple. Votre **fichier .eslintrc.json** peut avoir des propriétés supplémentaires et des membres supplémentaires des deux tableaux.

   ```json
   {
     "plugins": [
       "office-addins"
     ],
     "extends": [
       "plugin:office-addins/recommended"
     ]
   }
   ```

1. À la racine du projet, ouvrez le **fichier package.json** et assurez-vous que le tableau `scripts` possède le membre suivant.

   ```json
   "lint": "office-addin-lint check",
   ```

1. Exécutez le linter avec la commande suivante dans le terminal d’un éditeur, par exemple Visual Studio Code, ou dans une invite de commandes. Les problèmes trouvés par le linter apparaissent dans le terminal ou l’invite, et apparaissent également directement dans le code lorsque vous utilisez un éditeur qui prend en charge les messages linter, tels que Visual Studio Code.

   ```command&nbsp;line
   npm run lint
   ```

## <a name="next-steps"></a>Prochaines étapes

Essayez de créer votre propre Script Lab pour essayer des exemples intégrés.

### <a name="create-an-office-add-in"></a>Créer un complément Office

Vous pouvez créer rapidement un complément de base pour Excel, OneNote, Outlook, PowerPoint, Project ou Word en effectuant un [démarrage rapide de 5 minutes](../index.yml). Si vous avez déjà effectué un démarrage rapide et que vous voulez créer un complément légèrement plus complexe, vous devez essayer le [Didacticiel](../index.yml).

### <a name="explore-the-apis-with-script-lab"></a>Explorez des API avec Script Lab

Explorez la bibliothèque d’exemples intégrés dans [Script Lab](explore-with-script-lab.md) pour avoir une idée des capacités des API JavaScript Office.

## <a name="see-also"></a>Voir aussi

- [Concepts de base pour les compléments Office](../overview/core-concepts-office-add-ins.md)
- [Développement de compléments Office](../develop/develop-overview.md)
- [Concevoir des compléments Office](../design/add-in-design.md)
- [Test et débogage de compléments Office](../testing/test-debug-office-add-ins.md)
- [Publier des compléments Office](../publish/publish.md)
- [Découvrez le programme pour les développeurs Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program)