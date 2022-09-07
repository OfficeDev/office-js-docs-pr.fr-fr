---
title: Configuration de votre environnement de développement
description: Configurez votre environnement de développeur pour créer des compléments Office.
ms.date: 09/01/2022
ms.localizationpriority: medium
ms.openlocfilehash: 4e03ea7f55786107354f9d5a92e0cb30ffb559ec
ms.sourcegitcommit: 889d23061a9413deebf9092d675655f13704c727
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/07/2022
ms.locfileid: "67616002"
---
# <a name="set-up-your-development-environment"></a>Configuration de votre environnement de développement

Ce guide vous aide à configurer des outils afin de pouvoir créer des compléments Office en suivant nos guides de démarrage rapide ou nos didacticiels. Si vous les avez déjà installés, vous êtes prêt à commencer rapidement, par exemple [excel React démarrage rapide](../quickstarts/excel-quickstart-react.md).

## <a name="get-microsoft-365"></a>Obtenir Microsoft 365

Vous avez besoin d’un compte Microsoft 365. Vous pouvez bénéficier d’un abonnement Microsoft 365 gratuit de 90 jours renouvelable qui inclut toutes les applications Office en rejoignant le [programme de développement Microsoft 365](https://developer.microsoft.com/office/dev-program).

## <a name="install-the-environment"></a>Installer l’environnement

Il existe deux types d’environnements de développement parmi lesquels choisir. La structure des projets de complément Office créés dans les deux environnements est différente. Par conséquent, si plusieurs personnes travaillent sur un projet de complément, elles doivent toutes utiliser le même environnement. 

- **environnementNode.js** : recommandé. Dans cet environnement, vos outils sont installés et exécutés sur une ligne de commande. Le côté serveur de la partie application web du complément est écrit en JavaScript ou TypeScript et est hébergé dans un runtime Node.js. Il existe de nombreux outils de développement de compléments utiles dans cet environnement, tels qu’un linter Office et un planificateur/exécuteur de tâches appelé WebPack. L’outil de création et de génération de modèles automatiques de projet, Yo Office, est fréquemment mis à jour.
- **Environnement Visual Studio** : choisissez cet environnement uniquement si votre ordinateur de développement est Windows et que vous souhaitez développer le côté serveur du complément avec un langage et une infrastructure .NET, tels que ASP.NET. Les modèles de projet de complément dans Visual Studio ne sont pas mis à jour aussi fréquemment que ceux de l’environnement Node.js. Le code côté client ne peut pas être débogué avec le débogueur Visual Studio intégré, mais vous pouvez déboguer du code côté client avec les outils de développement de votre navigateur. Plus d’informations plus loin sous l’onglet **Environnement de Visual Studio** .

> [!NOTE]
> Visual Studio pour Mac n’inclut pas les modèles de structure de projet pour les compléments Office. Par conséquent, si votre ordinateur de développement est un Mac, vous devez utiliser l’environnement Node.js.

Sélectionnez l’onglet de l’environnement que vous choisissez. 

# <a name="nodejs-environment"></a>[ environnementNode.js](#tab/yeomangenerator)

Les principaux outils à installer sont les suivants :

- Node.js
- npm
- Éditeur de code de votre choix
- Yo Office
- Linter JavaScript Office

Ce guide part du principe que vous savez comment utiliser un outil en ligne de commande.

### <a name="install-nodejs-and-npm"></a>Installer Node.js et npm

Node.js est un runtime JavaScript que vous utilisez pour développer des compléments Office modernes.

Installez Node.js [en téléchargeant la dernière version recommandée à partir de son site web](https://nodejs.org). Suivez les instructions d’installation de votre système d’exploitation.

npm est un registre de logiciels open source à partir duquel télécharger les packages utilisés dans le développement de compléments Office. Il est généralement installé automatiquement lorsque vous installez Node.js. Pour vérifier si npm est déjà installé et voir la version installée, exécutez ce qui suit dans la ligne de commande.

```command&nbsp;line
npm -v
```

Si, pour une raison quelconque, vous souhaitez l’installer manuellement, exécutez ce qui suit dans la ligne de commande.

```command&nbsp;line
npm install npm -g
```

> [!TIP]
> Vous pouvez utiliser un gestionnaire de versions de nœud pour vous permettre de basculer entre plusieurs versions de Node.js et npm, mais cela n’est pas strictement nécessaire. Pour plus d’informations sur la procédure à suivre, [consultez les instructions de npm](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm).

### <a name="install-a-code-editor"></a>Installer un éditeur de code

Vous pouvez utiliser n’importe quel éditeur de code ou IDE qui prend en charge le développement côté client pour créer votre composant WebPart, par exemple :

- [Visual Studio Code](https://code.visualstudio.com/) (recommandé)
- [Atom](https://atom.io)
- [Webstorm](https://www.jetbrains.com/webstorm)

### <a name="install-the-yeoman-generator-mdash-yo-office"></a>Installer le générateur &mdash; Yeoman Yo Office

L’outil de création et de génération de modèles automatiques de projet est le [générateur Yeoman pour les compléments Office](../develop/yeoman-generator-overview.md), communément appelé **Yo Office**. Vous devez installer la dernière version de [Yeoman](https://github.com/yeoman/yo) et Yo Office. Pour installer ces outils globalement, exécutez la commande suivante via l’invite de commandes.

  ```command&nbsp;line
  npm install -g yo generator-office
  ```

### <a name="install-and-use-the-office-javascript-linter"></a>Installer et utiliser le linter JavaScript Office

Microsoft fournit un linter JavaScript pour vous aider à intercepter les erreurs courantes lors de l’utilisation de la bibliothèque JavaScript Office. Pour installer le linter, exécutez les deux commandes suivantes (après avoir [installé Node.js et npm](#install-nodejs-and-npm)).

```command&nbsp;line
npm install office-addin-lint --save-dev
npm install eslint-plugin-office-addins --save-dev
```

Si vous créez un projet de complément Office avec l’outil [générateur Yeoman pour compléments Office](../develop/yeoman-generator-overview.md) , le reste de la configuration est effectué pour vous. Exécutez le linter avec la commande suivante dans le terminal d’un éditeur, tel que Visual Studio Code, ou dans une invite de commandes. Les problèmes détectés par le linter apparaissent dans le terminal ou l’invite, et apparaissent également directement dans le code lorsque vous utilisez un éditeur qui prend en charge les messages linter, tels que Visual Studio Code. (Pour plus d’informations sur l’installation du générateur Yeoman, consultez [yeoman generator for Office Add-ins](../develop/yeoman-generator-overview.md).)

```command&nbsp;line
npm run lint
```

Si votre projet de complément a été créé d’une autre façon, procédez comme suit.

1. À la racine du projet, créez un fichier texte nommé **.eslintrc.json**, s’il n’en existe pas déjà un. Assurez-vous qu’il a des propriétés nommées `plugins` et `extends`, les deux, de tableau de type. Le `plugins` tableau doit inclure `"office-addins"` et le `extends` tableau doit inclure `"plugin:office-addins/recommended"`. Voici un exemple simple. Votre fichier **.eslintrc.json** peut avoir des propriétés supplémentaires et des membres supplémentaires des deux tableaux.

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

1. À la racine du projet, ouvrez le fichier **package.json** et assurez-vous que le `scripts` tableau a le membre suivant.

   ```json
   "lint": "office-addin-lint check",
   ```

1. Exécutez le linter avec la commande suivante dans le terminal d’un éditeur, tel que Visual Studio Code, ou dans une invite de commandes. Les problèmes détectés par le linter apparaissent dans le terminal ou l’invite, et apparaissent également directement dans le code lorsque vous utilisez un éditeur qui prend en charge les messages linter, tels que Visual Studio Code.

   ```command&nbsp;line
   npm run lint
   ```

# <a name="visual-studio-environment"></a>[Environnement Visual Studio](#tab/visualstudio)

### <a name="install-visual-studio"></a>Installer Visual Studio

Si Vous n’avez pas installé Visual Studio 2017 (pour Windows) ou version ultérieure, installez la dernière version à partir de [Visual Studio Downloads](https://visualstudio.microsoft.com/downloads/). Veillez à inclure la charge de **travail de développement Office/SharePoint** lorsque le programme d’installation vous demande de spécifier des charges de travail. Les autres charges de travail dont vous pouvez avoir besoin sont les **outils de développement Web pour** la **prise en charge du langage .NET, JavaScript et TypeScript** (pour le codage côté client du complément) et les charges de travail liées à la ASP.NET.

> [!TIP]
> Depuis juin 2022, les schémas XML du manifeste de complément Office installés avec Visual Studio ne sont pas la dernière version. Cela peut affecter les compléments, selon les fonctionnalités de complément qu’ils utilisent. Vous devrez peut-être mettre à jour les schémas XML pour le manifeste. Pour plus d’informations, consultez [Les erreurs de validation de schéma de manifeste dans les projets Visual Studio](../testing/troubleshoot-development-errors.md#manifest-schema-validation-errors-in-visual-studio-projects).

> [!NOTE]
> Pour plus d’informations sur le débogage du code côté client lorsque vous utilisez l’environnement Visual Studio, consultez [Déboguer des compléments Office dans Visual Studio](../develop/debug-office-add-ins-in-visual-studio.md). Déboguez le code côté serveur de la même façon que n’importe quelle application web créée dans Visual Studio. Voir [côté client ou côté serveur](../testing/debug-add-ins-overview.md#server-side-or-client-side).

---

## <a name="install-script-lab"></a>Installer Script Lab

Script Lab est un outil de prototypage rapide de code qui appelle les API de bibliothèque JavaScript Office. Script Lab est lui-même un complément Office et peut être installé à partir d’AppSource à [Script Lab](https://appsource.microsoft.com/marketplace/apps?search=script%20lab&page=1). Il existe une version pour Excel, PowerPoint et Word, et une version distincte pour Outlook. Pour plus d’informations sur l’utilisation de Script Lab, consultez [Explorer l’API JavaScript Office à l’aide de Script Lab](explore-with-script-lab.md).

## <a name="next-steps"></a>Prochaines étapes

Essayez de créer votre propre complément ou utilisez [Script Lab](explore-with-script-lab.md) pour essayer des exemples intégrés.

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