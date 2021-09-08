---
title: Configuration de votre environnement de développement
description: Configurer votre environnement de développement pour créer des Office de développement.
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: efc89b728117e2888cdebd2c5a132047fe662915
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58938554"
---
# <a name="set-up-your-development-environment"></a>Configuration de votre environnement de développement

Ce guide vous aide à configurer des outils pour créer des Office en suivant nos démarrages rapides ou didacticiels. Vous devez installer les outils à partir de la liste ci-dessous. Si vous avez déjà installé ces éléments, vous êtes prêt à commencer un démarrage rapide, tel que [celui-ci Excel React démarrage rapide.](../quickstarts/excel-quickstart-react.md)

- Node.js
- npm
- Un Microsoft 365 qui inclut la version d’abonnement de Office
- Éditeur de code de votre choix

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