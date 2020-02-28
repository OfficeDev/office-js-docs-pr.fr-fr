---
ms.date: 02/20/2020
title: Configurez votre complément Excel pour partager le runtime du navigateur (préversion)
ms.prod: excel
description: Configurez votre complément Excel pour partager le runtime du navigateur et exécuter le ruban, le volet des tâches et le code de fonction personnalisée dans le même runtime.
localization_priority: Priority
ms.openlocfilehash: 7945bd8fdb29a9d6d44d7d29676410a54bacf83f
ms.sourcegitcommit: dd6d00202f6466c27418247dad7bd136555a6036
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/26/2020
ms.locfileid: "42284134"
---
# <a name="configure-your-excel-add-in-to-use-a-shared-javascript-runtime-preview"></a>Configurez votre complément Excel pour utiliser un runtime JavaScript partagé (préversion).

[!include[Running custom functions in a shared runtime note](../includes/excel-shared-runtime-preview-note.md)]

Lors de l’exécution d’Excel sur Windows ou Mac, votre complément exécute le code des boutons du ruban, des fonctions personnalisées et du volet des tâches dans des environnements runtime JavaScript distincts. Cela permet de créer des limitations, telles que l'impossibilité de partager aisément des données globales ou de pouvoir accéder à l'ensemble des fonctionnalités CORS à partir d’une fonction personnalisée.

Vous pouvez toutefois configurer votre complément Excel pour partager un code dans un runtime JavaScript partagé. Vous pouvez ainsi améliorer la coordination dans votre complément et accéder aux DOM et CORS à partir de toutes les parties de votre complément. Il vous permet également d’exécuter un code lorsque le document s’ouvre ou pendant la fermeture du volet des tâches. Si vous voulez configurer votre complément pour utiliser un runtime partagé, suivez les instructions contenues dans cet article.

## <a name="create-the-add-in-project"></a>Création du projet de complément

Si vous démarrez un nouveau projet, suivez ces étapes pour utiliser le générateur Yeoman et créer le projet de complément Excel. Exécutez la commande suivante, puis répondez aux invites avec les réponses suivantes :

```command line
yo office
```

- Choose a project type (Choisissez un type de projet) : **projet de complément Fonctions personnalisées Excel**
- Choose a script type (Choisissez un type de script) :  **JavaScript**
- What do you want to name your add-in? (Comment souhaitez-vous nommer votre complément ?)  **My Office Add-in**

![Capture d’écran de réponse aux invites à partir d’Office pour créer le projet de complément.](../images/yo-office-excel-project.png)

Après avoir exécuté l’Assistant, le générateur crée le projet et installe les composants Node de prise en charge.

## <a name="configure-the-manifest"></a>Configurer le manifeste

Procédez comme suit pour configurer un projet nouveau ou existant de manière à utiliser un runtime partagé.

1. Démarrez Visual Studio Code et ouvrez le projet **My Office Add-in**.
2. Ouvrez le fichier **manifest.xml**.
3. Recherchez la section `<VersionOverrides>`, puis ajoutez l'exemple d'entrée suivante à la section `<Runtimes>`. La durée de vie doit être **longue** afin que les fonctions personnalisées puissent continuer de fonctionner même quand le volet Office est fermé. L'ID de ressources est `ContosoAddin.Url`, faisant par la suite référence à une chaîne dans la section des ressources. Vous pouvez utiliser n’importe quelle valeur d'ID de ressources souhaitée. Elle doit cependant correspondre à l'ID de ressources des autres éléments contenus dans les parties de votre complément.

   ```xml
   <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
     <Hosts>
       <Host xsi:type="Workbook">
       <Runtimes>
         <Runtime resid="ContosoAddin.Url" lifetime="long" />
       </Runtimes>
       <AllFormFactors>
   ```

4. Dans l’élément `<Page>`, remplacez l’emplacement de la source **Functions.Page.Url** par **ContosoAddin.Url**. Cet ID de ressources correspond à l'ID de ressources de `<Runtime>`. Veuillez noter que si vous ne disposez pas de fonctions personnalisées, vous n’aurez pas d'entrée de **Page**. Vous pouvez par conséquent ignorer cette étape.

   ```xml
   <AllFormFactors>
   ...
   <Page>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Page>
   ...
   ```

5. Dans la section `<DesktopFormFactor>`, changez la valeur **FunctionFile** de **Command.Url** pour utiliser **ContosoAddin.Url**. Veuillez noter que si vous n'avez pas de commandes d'action, vous ne disposerez pas d'entrée **FunctionFile**. Vous pouvez par conséquent ignorer cette étape.

   ```xml
   <DesktopFormFactor>
   <GetStarted>
   ...
   </GetStarted>
   <FunctionFile resid="ContosoAddin.Url"/>
   ```

6. Dans la section `<Action>`, remplacez l’emplacement de la source **Taskpane.Url** par **ContosoAddin.Url**. Veuillez noter que si vous n'avez pas de volet des tâches, vous ne disposerez pas de l'action **ShowTaskPane**. Vous pouvez par conséquent ignorer cette étape.

   ```xml
   <Action xsi:type="ShowTaskpane">
   <TaskpaneId>ButtonId1</TaskpaneId>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Action>
   ```

7. Ajoutez un nouvel **ID d’URL** pour **ContosoAddin.Url** pointant vers **taskpane.html**.

   ```xml
   <bt:Urls>
   <bt:Url id="Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js"/>
   ...
   <bt:Url id="ContosoAddin.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
   ...
   ```

8. Enregistrez vos changements et reconstruisez le projet.

   ```command line
   npm run build
   ```

## <a name="runtime-lifetime"></a>Durée de vie de l’exécution

Lorsque vous ajoutez l’élément `Runtime`, vous spécifiez également une durée de vie ayant une valeur de `long` ou de `short`. Configurez cette valeur sur `long` pour tirer parti de fonctionnalités telles que le démarrage de votre complément lorsque le document s’ouvre, continuer à exécuter un code après la fermeture du volet des tâches, ou utiliser CORS et DOM à partir de fonctions personnalisées.

Si vous configurez cette valeur sur `short`, le complément se comportera comme le comportement par défaut. Le complément démarre lorsque l’un des boutons de votre ruban est pressé, mais il peut se fermer lorsque l’exécution de votre gestionnaire de ruban se termine. De la même façon, le complément démarre lorsque le volet des tâches est ouvert, mais il se peut se fermer à la fermeture du volet des tâches.

```xml
<Runtimes>
  <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

## <a name="multiple-task-panes"></a>Multiples volets des tâches

Ne concevez pas votre complément pour utiliser plusieurs volets des tâches si vous envisagez d’utiliser le runtime partagé. Le runtime partagé prend uniquement en charge l’utilisation d’un volet des tâches. Notez que tout volet des tâches sans `<TaskpaneID>` est considéré comme un volet des tâches différent.

## <a name="next-steps"></a>Étapes suivantes

Essayez à présent des fonctionnalités du runtime partagé en consultant les articles suivants.

- [Appeler des API Excel à partir d'une fonction personnalisée](call-excel-apis-from-custom-function.md)

## <a name="see-also"></a>Voir aussi

- [Vue d’ensemble : exécutez votre code de complément dans un runtime JavaScript partagé (préversion)](custom-functions-shared-overview.md)
