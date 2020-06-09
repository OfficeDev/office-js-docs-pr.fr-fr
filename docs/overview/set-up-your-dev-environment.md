---
title: Configuration de votre environnement de développement
description: Configuration de votre environnement de développement pour créer des compléments Office
ms.date: 04/03/2020
localization_priority: Normal
ms.openlocfilehash: 6c3f533b56cafc8300837cc835b26361490afedb
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611953"
---
# <a name="set-up-your-development-environment"></a>Configuration de votre environnement de développement

Ce guide vous aide à configurer les outils de manière à pouvoir créer des compléments Office en suivant nos guides de démarrage rapide ou nos didacticiels. Vous devrez installer les outils à partir de la liste ci-dessous. Si ces éléments sont déjà installés, vous êtes prêt à commencer un démarrage rapide, tel que le [démarrage rapide de Microsoft Excel REACT](../quickstarts/excel-quickstart-react.md).

- Node.js
- npm
- Un compte Office 365 (la version d’abonnement d’Office)
- Un éditeur de code de votre choix

Ce guide suppose que vous sachiez comment utiliser un outil de ligne de commande. 

## <a name="install-nodejs"></a>Installer Node.js.

Node. js est un Runtime JavaScript dont vous aurez besoin pour développer des compléments Office modernes.

Installez node. js en [téléchargeant la version recommandée la plus récente à partir de leur site Web](https://nodejs.org). Suivez les instructions d’installation pour votre système d’exploitation.

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

Vous souhaiterez peut-être utiliser un gestionnaire de version de nœud pour vous permettre de basculer entre plusieurs versions de node. js et NPM, mais ce n’est pas obligatoire. Pour plus d’informations sur la procédure à suivre, [consultez les instructions de NPM](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm).

## <a name="get-office-365"></a>Obtenir Office 365

Si vous n’avez pas un compte Office 365, vous pouvez en obtenir un abonnement Office 365 gratuit et renouvelable de 90 jours en rejoignant le [Programme pour les développeurs Office 365](https://developer.microsoft.com/office/dev-program).

## <a name="install-a-code-editor"></a>Installer un éditeur de code

Vous pouvez utiliser n’importe quel éditeur de code ou IDE qui prend en charge le développement côté client pour créer votre composant WebPart, par exemple :

- [Visual Studio Code](https://code.visualstudio.com/)
- [Atom](https://atom.io)
- [Webstorm](https://www.jetbrains.com/webstorm)

## <a name="next-steps"></a>Étapes suivantes

Essayez de créer votre propre complément ou utilisez script Lab pour essayer des exemples intégrés.

### <a name="create-an-office-add-in"></a>Créer un complément Office

Vous pouvez créer rapidement un complément de base pour Excel, OneNote, Outlook, PowerPoint, Project ou Word en effectuant un [démarrage rapide de 5 minutes](../index.md). Si vous avez déjà effectué un démarrage rapide et que vous voulez créer un complément légèrement plus complexe, vous devez essayer le [Didacticiel](../index.md).

### <a name="explore-the-apis-with-script-lab"></a>Explorez des API avec Script Lab

Explorez la bibliothèque d’exemples intégrés dans [Script Lab](explore-with-script-lab.md) pour avoir une idée des capacités des API JavaScript Office.

## <a name="see-also"></a>Voir aussi

- [Création de compléments Office](../overview/office-add-ins-fundamentals.md)
- [Concepts de base pour les compléments Office](../overview/core-concepts-office-add-ins.md)
- [Développement de compléments Office](../develop/develop-overview.md)
- [Concevoir des compléments Office](../design/add-in-design.md)
- [Test et débogage de compléments Office](../testing/test-debug-office-add-ins.md)
- [Publier des compléments Office](../publish/publish.md)
