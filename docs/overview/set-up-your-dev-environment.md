---
title: Configuration de votre environnement de développement
description: Configurez votre environnement de développement pour créer des compléments Office.
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: 644194d7d0da479b13ac09d7e830af53e9a9838e
ms.sourcegitcommit: 42e6cfe51d99d4f3f05a3245829d764b28c46bbb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/23/2020
ms.locfileid: "48740832"
---
# <a name="set-up-your-development-environment"></a>Configuration de votre environnement de développement

Ce guide vous aide à configurer les outils de manière à pouvoir créer des compléments Office en suivant nos guides de démarrage rapide ou nos didacticiels. Vous devrez installer les outils à partir de la liste ci-dessous. Si ces éléments sont déjà installés, vous êtes prêt à commencer un démarrage rapide, tel que le [démarrage rapide de Microsoft Excel REACT](../quickstarts/excel-quickstart-react.md).

- Node.js
- npm
- Un compte Microsoft 365 qui inclut la version d’abonnement d’Office
- Un éditeur de code de votre choix

Ce guide suppose que vous sachiez comment utiliser un outil de ligne de commande. 

## <a name="install-nodejs"></a>Installer Node.js.

Node.js est un Runtime JavaScript dont vous aurez besoin pour développer des compléments Office modernes.

Installez Node.js en [téléchargeant la version recommandée la plus récente à partir de leur site Web](https://nodejs.org). Suivez les instructions d’installation pour votre système d’exploitation.

## <a name="install-npm"></a>Installer NPM

NPM est un registre de logiciels open source à partir duquel télécharger les packages utilisés dans le développement des compléments Office.

Pour installer NPM, exécutez ce qui suit dans la ligne de commande :

```command&nbsp;line
    npm install npm -g
```

Pour vérifier si NPM est déjà installé et voir la version installée, exécutez la commande suivante dans la ligne de commande :

```command&nbsp;line
npm -v
```

Vous souhaiterez peut-être utiliser un gestionnaire de version de nœud pour vous permettre de basculer entre plusieurs versions de Node.js et NPM, mais ce n’est pas obligatoire. Pour plus d’informations sur la procédure à suivre, [consultez les instructions de NPM](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm).

## <a name="get-office-365"></a>Obtenir Office 365

Si vous n’avez pas déjà un compte Office 365, vous pouvez obtenir gratuitement un abonnement de 90 jours renouvelable de Microsoft 365 en rejoignant le [Programme pour les développeurs Microsoft 365](https://developer.microsoft.com/office/dev-program).

## <a name="install-a-code-editor"></a>Installer un éditeur de code

Vous pouvez utiliser n’importe quel éditeur de code ou IDE qui prend en charge le développement côté client pour créer votre composant WebPart, par exemple :

- [Visual Studio Code](https://code.visualstudio.com/)
- [Atom](https://atom.io)
- [Webstorm](https://www.jetbrains.com/webstorm)

## <a name="next-steps"></a>Étapes suivantes

Essayez de créer votre propre complément ou utilisez script Lab pour essayer des exemples intégrés.

### <a name="create-an-office-add-in"></a>Créer un complément Office

Vous pouvez créer rapidement un complément de base pour Excel, OneNote, Outlook, PowerPoint, Project ou Word en effectuant un [démarrage rapide de 5 minutes](/office/dev/add-ins/). Si vous avez déjà effectué un démarrage rapide et que vous voulez créer un complément légèrement plus complexe, vous devez essayer le [Didacticiel](/office/dev/add-ins/).

### <a name="explore-the-apis-with-script-lab"></a>Explorez des API avec Script Lab

Explorez la bibliothèque d’exemples intégrés dans [Script Lab](explore-with-script-lab.md) pour avoir une idée des capacités des API JavaScript Office.

## <a name="see-also"></a>Voir aussi

- [Concepts de base pour les compléments Office](../overview/core-concepts-office-add-ins.md)
- [Développement de compléments Office](../develop/develop-overview.md)
- [Concevoir des compléments Office](../design/add-in-design.md)
- [Test et débogage de compléments Office](../testing/test-debug-office-add-ins.md)
- [Publier des compléments Office](../publish/publish.md)
- [En savoir plus sur le programme de développement Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program)
