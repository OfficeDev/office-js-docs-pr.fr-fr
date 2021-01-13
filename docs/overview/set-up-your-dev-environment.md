---
title: Configuration de votre environnement de développement
description: Configurer votre environnement de développement pour créer des add-ins Office.
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: eddf8bdf7b20a54667e6f8eb38bdace801ea1813
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839711"
---
# <a name="set-up-your-development-environment"></a>Configuration de votre environnement de développement

Ce guide vous aide à configurer des outils pour vous aider à créer des add-ins Office en suivant nos démarrages rapides ou didacticiels. Vous devez installer les outils à partir de la liste ci-dessous. Si vous avez déjà installé ces éléments, vous êtes prêt à commencer un démarrage rapide, tel que ce démarrage rapide [Excel React.](../quickstarts/excel-quickstart-react.md)

- Node.js
- npm
- Un compte Microsoft 365 qui inclut la version d’abonnement d’Office
- Éditeur de code de votre choix

Ce guide suppose que vous savez utiliser un outil de ligne de commande. 

## <a name="install-nodejs"></a>Installer Node.js.

Node.js est un runtime JavaScript dont vous aurez besoin pour développer des add-ins Office modernes.

Installez Node.js en [téléchargeant la dernière version recommandée à partir de leur site web.](https://nodejs.org) Suivez les instructions d’installation de votre système d’exploitation.

## <a name="install-npm"></a>Installer npm

npm est un registre de logiciel open source à partir duquel télécharger les packages utilisés dans le développement de modules office.

Pour installer npm, exécutez la commande suivante dans la ligne de commande :

```command&nbsp;line
    npm install npm -g
```

Pour vérifier si npm est déjà installé et voir la version installée, exécutez la commande suivante dans la ligne de commande :

```command&nbsp;line
npm -v
```

Vous pouvez utiliser un gestionnaire de version Node pour vous permettre de basculer entre plusieurs versions de Node.js et npm, mais cela n’est pas strictement nécessaire. Pour plus d’informations sur la façon de faire, voir [les instructions de npm.](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm)

## <a name="get-office-365"></a>Obtenir Office 365

Si vous n’avez pas déjà un compte Office 365, vous pouvez obtenir gratuitement un abonnement de 90 jours renouvelable de Microsoft 365 en rejoignant le [Programme pour les développeurs Microsoft 365](https://developer.microsoft.com/office/dev-program).

## <a name="install-a-code-editor"></a>Installer un éditeur de code

Vous pouvez utiliser n’importe quel éditeur de code ou IDE qui prend en charge le développement côté client pour créer votre composant WebPart, par exemple :

- [Visual Studio Code](https://code.visualstudio.com/)
- [Atom](https://atom.io)
- [Webstorm](https://www.jetbrains.com/webstorm)

## <a name="next-steps"></a>Étapes suivantes

Essayez de créer votre propre add-in ou utilisez Script Lab pour essayer des exemples intégrés.

### <a name="create-an-office-add-in"></a>Créer un complément Office

Vous pouvez créer rapidement un complément de base pour Excel, OneNote, Outlook, PowerPoint, Project ou Word en effectuant un [démarrage rapide de 5 minutes](../index.yml). Si vous avez déjà effectué un démarrage rapide et que vous voulez créer un complément légèrement plus complexe, vous devez essayer le [Didacticiel](../index.yml).

### <a name="explore-the-apis-with-script-lab"></a>Explorez des API avec Script Lab

Explorez la bibliothèque d’exemples intégrés dans [Script Lab](explore-with-script-lab.md) pour avoir une idée des capacités des API JavaScript Office.

## <a name="see-also"></a>Voir aussi

- [Concepts de base pour les compléments Office](../overview/core-concepts-office-add-ins.md)
- [Développement de add-ins Office](../develop/develop-overview.md)
- [Concevoir des compléments Office](../design/add-in-design.md)
- [Test et débogage de compléments Office](../testing/test-debug-office-add-ins.md)
- [Publier des compléments Office](../publish/publish.md)
- [Découvrez le programme pour les développeurs Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program)