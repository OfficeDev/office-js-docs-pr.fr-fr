---
title: Développement d’un complément Office avec Visual Studio Code
description: Comment développer un complément Office avec Visual Studio Code
ms.date: 01/16/2020
localization_priority: Priority
ms.openlocfilehash: 0aef01c5b892a0cf08254ca8ffd9dd751b993139
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608309"
---
# <a name="develop-office-add-ins-with-visual-studio-code"></a>Développement d’un complément Office avec Visual Studio Code

Cet article explique comment utiliser [Visual Studio Code (VS Code)](https://code.visualstudio.com) pour développer votre complément Office.

> [!NOTE]
> Pour en savoir plus sur l’utilisation de Visual Studio pour créer un complément Office, voir [Développer des compléments Office avec Visual Studio](develop-add-ins-visual-studio.md).

## <a name="prerequisites"></a>Conditions préalables

- [Visual Studio Code](https://code.visualstudio.com/)

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project-using-the-yeoman-generator"></a>Créez le projet de complément à l’aide du générateur Yeoman

Si vous utilisez le VS Code comme environnement de développement intégré (IDE), vous devez créer le projet de complément Office avec le [genérateur Yeoman pour les compléments Office](https://github.com/OfficeDev/generator-office). Le générateur Yeoman crée un projet Node js qui peut être géré avec VS Code ou n’importe quel autre éditeur. 

Pour créer un complément Office avec le générateur Yeoman, suivez les instructions dans le [démarrage rapide de 5 minutes](../index.md) qui correspond au type de complément que vous voulez créer.

## <a name="develop-the-add-in-using-vs-code"></a>Développer le complément à l’aide de VS Code

Lorsque le générateur Yeoman a terminé de créer le projet de complément, ouvrez le dossier racine du projet avec VS Code. 

> [!TIP]
> Dans Windows, vous pouvez accéder au répertoire racine du projet via la ligne de commande, puis entrer `code .` pour ouvrir ce dossier dans VS Code. Sur Mac, vous devez [ajouter la commande `code` au chemin d’accès](https://code.visualstudio.com/docs/setup/mac#_launching-from-the-command-line) avant de pouvoir utiliser cette commande pour ouvrir le dossier de projet dans VS Code.

Le générateur Yeoman crée un complément de base avec une fonctionnalité limitée. Vous pouvez personnaliser le complément en modifiant le [manifeste](add-in-manifests.md), HTML, JavaScript ou TypeScript et des fichiers CSS dans VS Code. Pour obtenir une description générale de la structure de projet et des fichiers dans le projet de complément que le générateur Yeoman crée, consultez les instructions du générateur Yeoman dans le [démarrage rapide de 5 minutes](../index.md) qui correspond au type de complément que vous avez créé.

## <a name="test-and-debug-the-add-in"></a>Tester et déboguer le complément

Les méthodes de test, de débogage et de résolution des problèmes liés aux compléments Office varient selon la plateforme. Pour plus d’informations, voir [Test et débogage de compléments Office](../testing/test-debug-office-add-ins.md).

## <a name="publish-the-add-in"></a>Publier le complément

[!include[instructions for publishing an Office Add-in](../includes/publish-add-in.md)]

## <a name="see-also"></a>Voir aussi

- [Création de compléments Office](../overview/office-add-ins-fundamentals.md)
- [Concepts de base pour les compléments Office](../overview/core-concepts-office-add-ins.md)
- [Développement de compléments Office](../develop/develop-overview.md)
- [Concevoir des compléments Office](../design/add-in-design.md)
- [Test et débogage de compléments Office](../testing/test-debug-office-add-ins.md)
- [Publish Office Add-ins](../publish/publish.md)