---
title: Créer des projets de complément Office à l’aide du générateur Yeoman
description: Découvrez comment créer des projets de complément Office à l’aide du générateur Yeoman pour les compléments Office.
ms.date: 08/19/2022
ms.localizationpriority: high
ms.openlocfilehash: f109c4dbc386a4cc23f72d0c67f9e4904360bba4
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/24/2022
ms.locfileid: "67422783"
---
# <a name="create-office-add-in-projects-using-the-yeoman-generator"></a>Créer des projets de complément Office à l’aide du générateur Yeoman

Le [générateur Yeoman pour compléments Office](https://github.com/OfficeDev/generator-office) (également appelé « Yo Office ») est un outil en ligne de commande interactif basé sur Node.js qui crée des projets de développement de compléments Office. Nous vous recommandons d’utiliser cet outil pour créer des projets de complément, sauf si vous souhaitez que le code côté serveur du complément soit dans un . Langage net (par exemple, C# ou VB.Net) ou vous souhaitez que le complément soit hébergé dans Internet Information Server (IIS). Dans l’une des deux dernières situations, [utilisez Visual Studio pour créer le complément](develop-add-ins-visual-studio.md).

Les projets créés par l’outil présentent les caractéristiques suivantes.

- Ils ont une configuration [npm](https://www.npmjs.com/) standard qui inclut un fichier **package.json** .
- Ils incluent plusieurs scripts utiles pour générer le projet, démarrer le serveur, charger le complément dans Office et d’autres tâches.
- Ils utilisent [webpack](https://webpack.js.org/) comme bundler et exécuteur de tâches de base.
- En mode de développement, ils sont hébergés sur localhost par le serveur webpack-dev-server basé sur Node.js de webpack, une version orientée développement du serveur [express](http://expressjs.com/) qui prend en charge le rechargement à chaud et le recompilation sur modification.
- Par défaut, toutes les dépendances sont installées par l’outil, mais vous pouvez reporter l’installation avec un argument de ligne de commande.
- Ils incluent un manifeste de complément complet.
- Ils disposent d’un complément de niveau « Hello World » qui est prêt à être exécuté dès que l’outil est terminé.
- Ils incluent un polyfill et un transpileur configuré pour transpiler TypeScript et les versions récentes de JavaScript en JavaScript ES5. Ces fonctionnalités garantissent que le complément est pris en charge dans tous les runtimes dans lesquels les compléments Office peuvent s’exécuter, y compris Internet Explorer.

> [!TIP]
> Si vous souhaitez vous écarter considérablement de ces choix, par exemple en utilisant un exécuteur de tâches différent ou un autre serveur, nous vous recommandons de choisir [l’option Manifeste uniquement](#manifest-only-option) lorsque vous exécutez l’outil.

## <a name="install-the-generator"></a>Installer le générateur

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="use-the-tool"></a>Utiliser l’outil

Démarrez l’outil avec la commande suivante dans une invite système (pas une fenêtre bash).

```command&nbsp;line
yo office 
```

Beaucoup de choses doivent être chargées, ce qui peut prendre 20 secondes avant le démarrage de l’outil. L’outil vous pose une série de questions. Pour certains, il vous suffit de taper une réponse à l’invite. Pour d’autres, vous avez une liste de réponses possibles. Si vous avez une liste donnée, sélectionnez-en une, puis entrez.

La première question vous demande de choisir entre six types de projets. Les options disponibles sont les suivantes :

- **Projet du volet Office Complément**
- **Projet de volet office de complément Office à l’aide de Angular framework**
- **Projet de volet office de complément Office à l’aide de React framework**
- **Projet du volet Office Complément pour les tâches qui prend en charge l’authentification unique**
- **Projet de complément Office contenant le manifeste uniquement**
- **Projet de complément Fonctions personnalisées Excel**

![Capture d’écran montrant l’invite de type de projet et les réponses possibles dans le générateur Yeoman.](../images/yo-office-project-type-prompt.png)

> [!NOTE]
> Le **projet du volet Office Complément Office qui prend en charge** l’option d’authentification unique produit un projet qui peut être utilisé pour voir comment fonctionne l’authentification unique (SSO) dans un complément. Le projet ne peut pas être utilisé comme base d’un complément de production. Pour obtenir un projet prenant en charge l’authentification unique qui peut être une base d’un complément de production, consultez la version « Complete » [de l’un des exemples d’authentification unique dans notre référentiel d’exemples](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth).
>
> Le **projet de complément Office contenant l’option manifeste uniquement** produit un projet qui contient un manifeste de complément de base et une structure minimale. Pour plus d’informations sur l’option, consultez [l’option Manifeste uniquement](#manifest-only-option).

La question suivante vous demande de choisir entre **TypeScript** et **JavaScript**. (Cette question est ignorée si vous avez choisi l’option manifeste uniquement dans la question précédente.)

![Capture d’écran montrant que l’utilisateur a choisi « Projet du volet office de complément Office » à la question précédente et affiche l’invite pour la langue, ainsi que les réponses possibles, TypeScript et JavaScript, dans le générateur Yeoman.](../images/yo-office-language-prompt.png)

Vous êtes ensuite invité à donner un nom au complément. Le nom que vous spécifiez sera utilisé dans le manifeste du complément, mais vous pourrez le modifier ultérieurement.

![Capture d’écran montrant que l’utilisateur a choisi TypeScript pour la question précédente et affiche l’invite pour le nom du complément dans le générateur Yeoman.](../images/yo-office-name-prompt.png)

Vous êtes ensuite invité à choisir l’application Office dans laquelle le complément doit s’exécuter. Vous pouvez choisir parmi six applications : **Excel**, **OneNote**, **Outlook**, **PowerPoint**, **Project** et **Word**. Vous devez en choisir un seul, mais vous pouvez modifier le manifeste ultérieurement pour prendre en charge les applications Office supplémentaires. L’exception est Outlook. Un manifeste qui prend en charge Outlook ne peut prendre en charge aucune autre application Office.

![Capture d’écran montrant que l’utilisateur a nommé le projet « Complément Contoso » et affiche l’invite d’application Office et les réponses possibles dans le générateur Yeoman.](../images/yo-office-host-prompt.png)

Une fois que vous avez répondu à cette question, le générateur crée le projet et installe les dépendances. Vous pouvez voir **des messages WARN** dans la sortie npm à l’écran. Vous pouvez les ignorer. Vous pouvez également voir des messages indiquant que des vulnérabilités ont été détectées. Vous pouvez les ignorer pour l’instant, mais vous devrez éventuellement les corriger avant que votre complément ne soit mis en production. Pour plus d’informations sur la résolution des vulnérabilités, ouvrez votre navigateur et recherchez « vulnérabilité npm ».

Si la création réussit, vous verrez une **félicitations !** dans la fenêtre de commande, suivi de quelques étapes suivantes suggérées. (Si vous utilisez le générateur dans le cadre d’un démarrage rapide ou d’un didacticiel, ignorez les étapes suivantes dans la fenêtre de commande et suivez les instructions de l’article.)

> [!TIP]
> Si vous souhaitez créer la structure d’un projet de complément Office, mais reporter l’installation des dépendances, ajoutez l’option `--skip-install` à la `yo office` commande. Voici un exemple de code.
>
> ```command&nbsp;line
> yo office --skip-install
> ```
>
> Lorsque vous êtes prêt à installer les dépendances, accédez au dossier racine du projet dans une invite de commandes et entrez `npm install`.

## <a name="manifest-only-option"></a>Option manifeste uniquement

Cette option crée uniquement un manifeste pour un complément. Le projet résultant n’a pas de complément Hello World, aucun des scripts ni aucune des dépendances. Utilisez cette option dans les scénarios suivants.

- Vous souhaitez utiliser différents outils que celui qu’un projet générateur Yeoman installe et configure par défaut. Par exemple, vous souhaitez utiliser un autre bundler, un transpileur, un exécuteur de tâches ou un serveur de développement différent.
- Vous souhaitez utiliser une infrastructure de développement d’applications web, autre que Angular ou React, telle que Vue.

Pour obtenir un exemple d’utilisation du générateur avec l’option manifeste uniquement, consultez [Utiliser Vue pour créer un complément du volet Office Excel](../quickstarts/excel-quickstart-vue.md).

## <a name="use-command-line-parameters"></a>Utiliser des paramètres de ligne de commande

Vous pouvez également ajouter des paramètres à la `yo office` commande. Les deux plus courantes sont :

- `yo office --details`: cela génère une brève aide sur tous les autres paramètres de ligne de commande.
- `yo office --skip-install`: cela empêche le générateur d’installer les dépendances.

Pour obtenir des informations détaillées sur les paramètres de ligne de commande, consultez le [lisez-moi du générateur dans yeoman generator for Office Add-ins](https://github.com/officedev/generator-office).

## <a name="troubleshooting"></a>Résolution des problèmes

Si vous rencontrez des problèmes à l’aide de l’outil, la première étape consiste à le réinstaller pour vous assurer que vous disposez de la dernière version. (Pour plus [d’informations, consultez Installer le générateur](#install-the-generator) .) Si cela ne résout pas le problème, recherchez [l’outil dans le référentiel GitHub pour](https://github.com/OfficeDev/generator-office/issues) voir si quelqu’un d’autre a rencontré le même problème et trouvé une solution. Si personne ne l’a fait, [créez un problème](https://github.com/OfficeDev/generator-office/issues/new?assignees=&labels=needs+triage&template=bug_report.md&title=).
